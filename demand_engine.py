"""Demand Management data-quality, forecast-flow and Excel-export engine.

This companion engine reads demand-planning CSV/TXT exports from one folder.  It
uses the same schema-tolerant conventions as twin_engine.py and can be run from
the command line or through demand_app.py.

Example:
python demand_engine.py --folder "C:\\path\\APP_OUTPUT_FILES" --output Demand_Management_Checks.xlsx
"""
from __future__ import annotations

import argparse
import io
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Mapping, Sequence

import numpy as np
import pandas as pd

from twin_engine import (
    APP_VERSION, EXCEL_MAX_DATA_ROWS, SourceTable, build_dfu_mapping_loss,
    build_forecast_flow, canonical_column, clean_text, dimension_frame,
    first_column, load_sources, month_start, normalise_name, numeric, parse_dates,
    quantity_mt,
)

DM_VERSION = "1.0.0"

ROLE_RULES = [
    ("DFU master", ("DFU_MASTER", "DFUMASTER")),
    ("DFU-to-SKU mapping", ("DFUTOSKU_FE", "DFUTOSKU_MAP")),
    ("BU forecast", ("DFUTOSKUFCST_BU", "BU_FCST", "BU_FORECAST")),
    ("SKU forecast", ("DFUTOSKUFCST", "SKU_FCST", "SKU_FORECAST")),
    ("History / actuals", ("HISTORY", "HIST", "ACTUAL", "SALES")),
    ("Customer orders", ("CUSTORDER", "CUSTOMER_ORDER", "ORDERS")),
    ("Net demand signal", ("INDDMDVIEW", "NETTED_FCST", "NET_DEMAND")),
    ("Order fulfilment", ("FULFIL", "SCHED_QTY", "SHIPMENT", "DELIVERY")),
    ("Other master", ("MASTER", "SKU", "ITEM")),
]

ROLE_KEYS = {
    "DFU master": ("DFU", "DMDGROUP", "PLANT"),
    "DFU-to-SKU mapping": ("DFU", "DMDGROUP", "PLANT", "ITEM"),
    "BU forecast": ("DFU", "DMDGROUP", "PLANT", "BU", "DATE"),
    "SKU forecast": ("ITEM", "PLANT", "DATE"),
    "History / actuals": ("ITEM", "PLANT", "DATE"),
    "Customer orders": ("ITEM", "PLANT", "DATE"),
    "Net demand signal": ("ITEM", "PLANT", "DATE"),
    "Order fulfilment": ("ITEM", "PLANT", "DATE"),
}


def role_for(source: SourceTable) -> str:
    n = source.name
    # The BU-specific file must be tested before generic DFUTOSKUFCST.
    for role, tokens in ROLE_RULES:
        if any(normalise_name(token) in n for token in tokens):
            if role == "SKU forecast" and "DFUTOSKUFCST_BU" in n:
                continue
            return role
    return "Unclassified"


@dataclass
class DemandModel:
    folder: Path
    simulation: str | None
    sources: dict[str, SourceTable]
    tables: dict[str, pd.DataFrame] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)
    generated_at: datetime = field(default_factory=datetime.now)

    def kpis(self) -> dict[str, float | int]:
        quality = self.tables.get("data_quality", pd.DataFrame())
        loss = self.tables.get("dfu_mapping_loss", pd.DataFrame())
        flow = self.tables.get("forecast_flow", pd.DataFrame())
        high = quality[quality.get("SEVERITY", pd.Series(dtype=str)).eq("High")] if not quality.empty else quality
        total = float(flow.get("TOTAL_FCST_MT", pd.Series(dtype=float)).sum())
        sched = float(flow.get("SCHED_QTY_MT", pd.Series(dtype=float)).sum())
        return {
            "Files": len(self.sources),
            "Rows": int(sum(len(s.data) for s in self.sources.values())),
            "Quality issues": int(len(quality)),
            "High issues": int(len(high)),
            "Mapping lost MT": float(loss.get("LOST_MT", pd.Series(dtype=float)).sum()),
            "Weighted MET %": sched / total * 100 if total else np.nan,
            "Where-to-hit actions": int(len(self.tables.get("where_to_hit", pd.DataFrame()))),
        }


def _field_col(df: pd.DataFrame, field: str) -> str | None:
    return canonical_column(df, field)


def build_file_inventory(sources: Mapping[str, SourceTable]) -> pd.DataFrame:
    rows = []
    for s in sources.values():
        date_col = _field_col(s.data, "DATE")
        qty_col = _field_col(s.data, "QTY")
        dates = parse_dates(s.data[date_col]) if date_col else pd.Series(dtype="datetime64[ns]")
        rows.append({
            "FILE": s.path.name, "ROLE": role_for(s), "ROWS": len(s.data), "COLUMNS": len(s.data.columns),
            "DATE_FIELD": date_col or "", "MIN_DATE": dates.min() if len(dates) else pd.NaT,
            "MAX_DATE": dates.max() if len(dates) else pd.NaT, "QUANTITY_FIELD": qty_col or "",
            "SIMULATION_FILTER": s.simulation_note, "COLUMN_LIST": ", ".join(s.data.columns),
        })
    return pd.DataFrame(rows).sort_values(["ROLE", "FILE"]).reset_index(drop=True)


def build_data_quality(sources: Mapping[str, SourceTable]) -> pd.DataFrame:
    rows: list[dict] = []
    def add(severity: str, category: str, source: SourceTable, field: str, count: int, rate: float, action: str) -> None:
        if count <= 0: return
        rows.append({"SEVERITY": severity, "CATEGORY": category, "FILE": source.path.name, "ROLE": role_for(source),
                     "FIELD": field, "COUNT": int(count), "RATE_PCT": float(rate), "ACTION": action})

    for source in sources.values():
        df = source.data
        role = role_for(source)
        total = max(len(df), 1)
        if df.empty:
            add("High", "Empty file", source, "", 1, 100.0, "Check extraction and simulation selection.")
            continue
        duplicates = int(df.duplicated().sum())
        add("Medium", "Duplicate rows", source, "All columns", duplicates, duplicates / total * 100,
            "Remove duplicate extracts or define the correct business key.")
        expected = ROLE_KEYS.get(role, ())
        for field in expected:
            col = _field_col(df, field)
            if not col:
                add("High", "Missing required column", source, field, len(df), 100.0,
                    f"Add/map a {field} column for {role} checks.")
                continue
            raw = clean_text(df[col], "")
            missing = int(raw.eq("").sum())
            add("High" if missing / total > .05 else "Medium", "Blank business key", source, col, missing, missing / total * 100,
                "Correct upstream master/transaction keys before planning.")
        date_col = _field_col(df, "DATE")
        if date_col:
            raw_present = clean_text(df[date_col], "").ne("")
            invalid = int((raw_present & parse_dates(df[date_col]).isna()).sum())
            add("High", "Invalid date", source, date_col, invalid, invalid / total * 100,
                "Standardise the export date format and reject invalid buckets.")
        qty_col = _field_col(df, "QTY")
        if qty_col:
            q = numeric(df[qty_col])
            negative = int(q.lt(0).sum())
            add("Medium", "Negative quantity", source, qty_col, negative, negative / total * 100,
                "Confirm whether negatives are valid returns/cancellations; otherwise correct source data.")
            zero = int(q.eq(0).sum())
            add("Low", "Zero quantity", source, qty_col, zero, zero / total * 100,
                "Remove non-contributing rows or confirm that zero is a required placeholder.")
        # Duplicate natural key is more useful than whole-row duplicates.
        natural_cols = [c for f in expected if (c := _field_col(df, f))]
        if natural_cols:
            dup_key = int(df.duplicated(natural_cols, keep=False).sum())
            add("High" if role in {"DFU master", "DFU-to-SKU mapping"} else "Medium", "Duplicate business key",
                source, ", ".join(natural_cols), dup_key, dup_key / total * 100,
                "Validate uniqueness or add the missing key dimension (type, customer or version).")
    columns = ["SEVERITY", "CATEGORY", "FILE", "ROLE", "FIELD", "COUNT", "RATE_PCT", "ACTION"]
    if not rows: return pd.DataFrame(columns=columns)
    order = pd.CategoricalDtype(["High", "Medium", "Low"], ordered=True)
    out = pd.DataFrame(rows); out["SEVERITY"] = out["SEVERITY"].astype(order)
    return out.sort_values(["SEVERITY", "COUNT"], ascending=[True, False]).reset_index(drop=True)


def _measure_fact(source: SourceTable, value_name: str) -> pd.DataFrame:
    df = source.data
    out = dimension_frame(df)
    date_col = canonical_column(df, "DATE")
    qty_col = canonical_column(df, "QTY")
    out["MONTH"] = month_start(parse_dates(df[date_col])) if date_col else pd.NaT
    out[value_name] = quantity_mt(df[qty_col], qty_col) if qty_col else 0.0
    return out


def _sources_for_role(sources: Mapping[str, SourceTable], role: str) -> list[SourceTable]:
    return [s for s in sources.values() if role_for(s) == role]


def build_forecast_vs_orders(sources: Mapping[str, SourceTable], flow: pd.DataFrame) -> pd.DataFrame:
    keys = ["MONTH", "BU", "PLANT", "ITEM"]
    if not flow.empty:
        cols = [*keys, "NETTED_FCST_MT", "CUSTORDER_MT", "TOTAL_FCST_MT", "SCHED_QTY_MT"]
        out = flow.groupby(keys, dropna=False, observed=True, as_index=False)[[c for c in cols[4:] if c in flow]].sum()
    else:
        out = pd.DataFrame(columns=keys)
    # Add actual/history when available.
    hist_frames = [_measure_fact(s, "ACTUAL_MT") for s in _sources_for_role(sources, "History / actuals")]
    if hist_frames:
        hist = pd.concat(hist_frames, ignore_index=True).groupby(keys, dropna=False, observed=True, as_index=False)["ACTUAL_MT"].sum()
        out = hist if out.empty else out.merge(hist, on=keys, how="outer")
    for col in ["NETTED_FCST_MT", "CUSTORDER_MT", "TOTAL_FCST_MT", "SCHED_QTY_MT", "ACTUAL_MT"]:
        if col not in out: out[col] = 0.0
        out[col] = out[col].fillna(0.0)
    out["ORDER_SHARE_PCT"] = np.where(out["TOTAL_FCST_MT"].gt(0), out["CUSTORDER_MT"] / out["TOTAL_FCST_MT"] * 100, np.nan)
    out["UNMET_MT"] = (out["TOTAL_FCST_MT"] - out["SCHED_QTY_MT"]).clip(lower=0)
    out["MET_PCT"] = np.where(out["TOTAL_FCST_MT"].gt(0), out["SCHED_QTY_MT"] / out["TOTAL_FCST_MT"] * 100, np.nan)
    return out.sort_values(["MONTH", "BU", "PLANT", "ITEM"]).reset_index(drop=True)


def build_bias_coverage(forecast_vs_orders: pd.DataFrame) -> pd.DataFrame:
    if forecast_vs_orders.empty or forecast_vs_orders["ACTUAL_MT"].abs().sum() == 0:
        return pd.DataFrame(columns=["BU", "PLANT", "ACTUAL_MT", "FORECAST_MT", "BIAS_MT", "BIAS_PCT", "WAPE_PCT", "ACTUAL_MONTHS", "FORECAST_MONTHS", "BUCKET_COVERAGE_PCT"])
    df = forecast_vs_orders.copy()
    df["FORECAST_MT"] = np.where(df["TOTAL_FCST_MT"].ne(0), df["TOTAL_FCST_MT"], df["NETTED_FCST_MT"])
    df["ERROR_MT"] = df["FORECAST_MT"] - df["ACTUAL_MT"]
    df["ABS_ERROR_MT"] = df["ERROR_MT"].abs()
    df["ACTUAL_BUCKET"] = df["ACTUAL_MT"].ne(0).astype(int)
    df["FORECAST_BUCKET"] = df["FORECAST_MT"].ne(0).astype(int)
    agg = (df.groupby(["BU", "PLANT"], dropna=False, observed=True)
           .agg(ACTUAL_MT=("ACTUAL_MT", "sum"), FORECAST_MT=("FORECAST_MT", "sum"), BIAS_MT=("ERROR_MT", "sum"),
                ABS_ERROR_MT=("ABS_ERROR_MT", "sum"), ACTUAL_MONTHS=("ACTUAL_BUCKET", "sum"), FORECAST_MONTHS=("FORECAST_BUCKET", "sum"))
           .reset_index())
    agg["BIAS_PCT"] = np.where(agg["ACTUAL_MT"].ne(0), agg["BIAS_MT"] / agg["ACTUAL_MT"] * 100, np.nan)
    agg["WAPE_PCT"] = np.where(agg["ACTUAL_MT"].abs().gt(0), agg["ABS_ERROR_MT"] / agg["ACTUAL_MT"].abs() * 100, np.nan)
    agg["BUCKET_COVERAGE_PCT"] = np.where(agg["ACTUAL_MONTHS"].gt(0), agg["FORECAST_MONTHS"] / agg["ACTUAL_MONTHS"] * 100, np.nan)
    return agg.drop(columns="ABS_ERROR_MT").sort_values("WAPE_PCT", ascending=False).reset_index(drop=True)


def build_where_to_hit(quality: pd.DataFrame, loss: pd.DataFrame, fvo: pd.DataFrame) -> pd.DataFrame:
    rows: list[dict] = []
    priority_map = {"High": 1, "Medium": 2, "Low": 3}
    for _, r in quality.head(100).iterrows():
        sev = str(r["SEVERITY"])
        rows.append({"PRIORITY": priority_map.get(sev, 3), "AREA": "Data quality", "WHERE": r["FILE"],
                     "ISSUE": f"{r['CATEGORY']} — {r['FIELD']}", "IMPACT": f"{int(r['COUNT']):,} rows ({r['RATE_PCT']:.1f}%)",
                     "ACTION": r["ACTION"]})
    if not loss.empty:
        location_col = "LOC" if "LOC" in loss else "LOCATION"
        reason_col = "REASON" if "REASON" in loss else "ISSUE"
        grouped = (loss.groupby(["BU", location_col, reason_col], dropna=False, observed=True, as_index=False)["LOST_MT"].sum()
                   .sort_values("LOST_MT", ascending=False).head(100))
        for _, r in grouped.iterrows():
            rows.append({"PRIORITY": 1, "AREA": "DFU mapping", "WHERE": f"{r['BU']} / {r[location_col]}",
                         "ISSUE": r[reason_col], "IMPACT": f"{r['LOST_MT']:,.1f} MT forecast lost",
                         "ACTION": "Create/correct DFUTOSKU_FE mapping, then rerun SKU forecast generation."})
    if not fvo.empty:
        gaps = (fvo.groupby(["BU", "PLANT"], dropna=False, observed=True, as_index=False)[["TOTAL_FCST_MT", "SCHED_QTY_MT", "UNMET_MT"]].sum()
                .sort_values("UNMET_MT", ascending=False).head(100))
        for _, r in gaps[gaps["UNMET_MT"] > 0].iterrows():
            pct = r["SCHED_QTY_MT"] / r["TOTAL_FCST_MT"] * 100 if r["TOTAL_FCST_MT"] else np.nan
            rows.append({"PRIORITY": 1 if pct < 80 else 2, "AREA": "Order fulfilment", "WHERE": f"{r['BU']} / {r['PLANT']}",
                         "ISSUE": f"Supply below demand; MET {pct:.1f}%", "IMPACT": f"{r['UNMET_MT']:,.1f} MT unmet",
                         "ACTION": "Review constrained items/resources, supply dates and order priorities."})
    columns = ["PRIORITY", "AREA", "WHERE", "ISSUE", "IMPACT", "ACTION"]
    if not rows: return pd.DataFrame(columns=columns)
    return pd.DataFrame(rows).sort_values(["PRIORITY", "AREA", "IMPACT"]).reset_index(drop=True)


def build_demand_model(folder: str | Path, simulation: str | None = None) -> DemandModel:
    sources, warnings = load_sources(folder, simulation)
    model = DemandModel(Path(folder), simulation, sources, warnings=warnings)
    inventory = build_file_inventory(sources)
    quality = build_data_quality(sources)
    flow = build_forecast_flow(sources)
    loss = build_dfu_mapping_loss(sources)
    fvo = build_forecast_vs_orders(sources, flow)
    bias = build_bias_coverage(fvo)
    hit = build_where_to_hit(quality, loss, fvo)
    model.tables.update({"file_inventory": inventory, "data_quality": quality, "forecast_flow": flow,
                         "forecast_vs_orders": fvo, "bias_coverage": bias, "dfu_mapping_loss": loss,
                         "where_to_hit": hit})
    roles = set(inventory["ROLE"])
    for required in ("DFU master", "DFU-to-SKU mapping", "BU forecast", "SKU forecast", "History / actuals", "Customer orders"):
        if required == "Customer orders" and not flow.empty and flow.get("CUSTORDER_MT", pd.Series(dtype=float)).abs().sum() > 0:
            continue
        if required not in roles: model.warnings.append(f"{required} export not identified; related checks are partial.")
    return model


SHEETS = [
    ("File Inventory", "file_inventory"), ("Data Quality", "data_quality"),
    ("Forecast Flow", "forecast_flow"), ("Forecast vs Orders", "forecast_vs_orders"),
    ("Bias Coverage", "bias_coverage"), ("DFU Mapping Loss", "dfu_mapping_loss"),
    ("Where to Hit", "where_to_hit"),
]


def _excel_safe(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy().replace([np.inf, -np.inf], np.nan)
    for c in out.columns:
        if isinstance(out[c].dtype, pd.CategoricalDtype): out[c] = out[c].astype(str)
        if pd.api.types.is_datetime64_any_dtype(out[c]):
            try: out[c] = out[c].dt.tz_localize(None)
            except (TypeError, AttributeError): pass
    return out


def demand_excel_bytes(model: DemandModel) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter", datetime_format="dd-mmm-yyyy", date_format="dd-mmm-yyyy") as writer:
        wb = writer.book
        navy, blue, pale, white, red, amber = "#17365D", "#1F4E78", "#D9EAF7", "#FFFFFF", "#C00000", "#F4B183"
        title = wb.add_format({"bold": True, "font_size": 21, "font_color": white, "bg_color": navy, "valign": "vcenter"})
        header = wb.add_format({"bold": True, "font_color": white, "bg_color": blue, "border": 1, "text_wrap": True})
        nav = wb.add_format({"font_color": blue, "underline": True, "bg_color": pale, "border": 1})
        label = wb.add_format({"bold": True, "font_color": navy, "bg_color": pale, "border": 1})
        value = wb.add_format({"bold": True, "font_size": 15, "font_color": navy, "border": 1, "num_format": "#,##0.0"})
        mt = wb.add_format({"num_format": "#,##0.0"})

        home = wb.add_worksheet("Home"); writer.sheets["Home"] = home
        home.set_column("A:A", 3); home.set_column("B:B", 30); home.set_column("C:C", 72); home.set_row(1, 34)
        home.merge_range("B2:C2", "Demand Management Checks", title)
        home.write("B3", f"Version {DM_VERSION} | Generated {model.generated_at:%d-%b-%Y %H:%M}")
        home.write("B5", "Folder", label); home.write("C5", str(model.folder))
        home.write("B6", "Simulation", label); home.write("C6", model.simulation or "All")
        home.write_row("B8", ["Sheet", "Purpose"], header)
        descriptions = {
            "File Inventory": "Master, transaction and output file coverage and date ranges",
            "Data Quality": "Missing fields, blank keys, invalid dates, duplicates and quantity checks",
            "Forecast Flow": "BU → SKU → netted forecast + orders → total → schedule → MET",
            "Forecast vs Orders": "Demand signal, orders, actuals, schedule and unmet MT",
            "Bias Coverage": "Bias, WAPE and bucket coverage where actuals are available",
            "DFU Mapping Loss": "DFU/month forecast that did not reach SKU forecast",
            "Where to Hit": "Prioritised action list",
        }
        r = 8
        for sheet, _ in SHEETS:
            home.write_url(r, 1, f"internal:'{sheet}'!A1", nav, sheet); home.write(r, 2, descriptions[sheet]); r += 1
        if model.warnings:
            r += 1; home.write(r, 1, "Coverage notes", header); r += 1
            for w in model.warnings: home.merge_range(r, 1, r, 2, w); r += 1

        summary = wb.add_worksheet("Executive Summary"); writer.sheets["Executive Summary"] = summary
        summary.set_column("A:A", 3); summary.set_column("B:H", 20); summary.set_row(1, 34)
        summary.merge_range("B2:H2", "Demand Management Executive Summary", title)
        kpis = model.kpis()
        for i, (k, v) in enumerate(kpis.items()):
            rr = 4 + (i // 4) * 3; cc = 1 + (i % 4) * 2
            summary.write(rr, cc, k, label); summary.write(rr+1, cc, v, value)
        summary.write_url("B12", "internal:'Home'!B2", nav, "← Home")
        hit = model.tables["where_to_hit"].head(15)
        if not hit.empty:
            hit.to_excel(writer, sheet_name="Executive Summary", startrow=14, startcol=1, index=False)
            for c, name in enumerate(hit.columns): summary.write(14, 1+c, name, header)

        for sheet, key in SHEETS:
            df = _excel_safe(model.tables[key])
            if len(df) > EXCEL_MAX_DATA_ROWS: df = df.iloc[:EXCEL_MAX_DATA_ROWS]
            df.to_excel(writer, sheet_name=sheet, startrow=3, index=False)
            ws = writer.sheets[sheet]; ws.freeze_panes(4, 1); ws.set_row(0, 30)
            ws.merge_range(0, 0, 0, max(2, len(df.columns)-1), sheet, title)
            ws.write_url(1, 0, "internal:'Home'!B2", nav, "← Home")
            for c, col in enumerate(df.columns):
                ws.write(3, c, col, header)
                width = min(48, max(11, len(str(col))+2, *(df[col].astype(str).head(200).str.len().tolist() or [0])))
                ws.set_column(c, c, width, mt if str(col).endswith("_MT") else None)
            if len(df) and len(df.columns):
                name = "T_" + "".join(ch for ch in sheet if ch.isalnum())
                ws.add_table(3, 0, len(df)+3, len(df.columns)-1, {"name": name, "style": "Table Style Medium 2", "columns": [{"header": str(c)} for c in df.columns]})
            if "SEVERITY" in df.columns and len(df):
                c = df.columns.get_loc("SEVERITY")
                ws.conditional_format(4, c, len(df)+3, c, {"type": "text", "criteria": "containing", "value": "High", "format": wb.add_format({"bg_color": "#FFC7CE", "font_color": red})})
            if "LOST_MT" in df.columns and len(df):
                c = df.columns.get_loc("LOST_MT"); ws.conditional_format(4, c, len(df)+3, c, {"type": "data_bar", "bar_color": amber})
        wb.set_properties({"title": "Demand Management Checks", "author": "Demand Management Digital Twin",
                           "comments": f"Uses twin_engine.py {APP_VERSION} and demand_engine.py {DM_VERSION}"})
    return bio.getvalue()


def write_demand_excel(
    model: DemandModel,
    output_path: str | Path | None = None
) -> Path:

    # Default output folder:
    # ESP_RUN_FILES\APP_OUTPUT_FILES
    output_folder = model.folder / "APP_OUTPUT_FILES"
    output_folder.mkdir(parents=True, exist_ok=True)

    # Timestamp avoids PermissionError when an older Excel file is open
    timestamp = pd.Timestamp.now().strftime("%d%b_%H%M")

    if output_path:
        supplied = Path(output_path).expanduser()

        # If only a filename was supplied, keep it inside APP_OUTPUT_FILES
        if supplied.parent == Path("."):
            path = output_folder / supplied.name
        else:
            path = supplied
            path.parent.mkdir(parents=True, exist_ok=True)
    else:
        path = output_folder / f"Demand_Management_Checks_{timestamp}.xlsx"

    path.write_bytes(demand_excel_bytes(model))

    return path


def main(argv: Sequence[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="Run Demand Management checks and create an Excel workbook.")
    p.add_argument("--folder", required=True, help="Folder containing demand-planning CSV/TXT exports")
    p.add_argument("--simulation", default=None, help="Optional simulation/scenario")
    p.add_argument("--output", default=None, help="Optional output .xlsx path. Default: ESP_RUN_FILES\\APP_OUTPUT_FILES")
    a = p.parse_args(argv)
    model = build_demand_model(a.folder, a.simulation)
    path = write_demand_excel(model, a.output)
    print(f"Created: {path.resolve()}")
    print(f"Loaded {len(model.sources)} files; {len(model.tables['data_quality'])} quality issue rows")
    for warning in model.warnings: print(f"WARNING: {warning}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
