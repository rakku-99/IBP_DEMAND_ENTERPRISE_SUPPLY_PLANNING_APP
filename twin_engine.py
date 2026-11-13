"""ESP Digital Twin data engine and Excel exporter.

Designed for Blue Yonder ESP CSV exports.  The loader is intentionally tolerant of
column-name differences between implementations: source columns are normalised to
UPPER_SNAKE_CASE, a schema report is retained, and common business fields are
resolved from aliases.

Command line example
--------------------
python twin_engine.py --folder "C:\\path\\ESP_RUN_FILES" \
    --simulation MONTHLY_WHATIF --output ESP_Network_Digital_Twin.xlsx

Quantities in source exports are assumed to be kg and are converted to MT by
/ 1000.  A source column explicitly containing "_MT" is treated as already MT.
Business exception: UDT_BU_HEAD_FCST.SBU_FCST is contractually already MT.
OPTIONSET is never used as a simulation column; in ESP exports it is a measure
name (for example SNOPWI_BATCH_STORE_RES_PROJ).
"""
from __future__ import annotations

import argparse
import io
import math
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence

import numpy as np
import pandas as pd

APP_VERSION = "2.5.0"
APP_OUTPUT_DIRNAME = "APP_OUTPUT_FILES"
EXCEL_MAX_DATA_ROWS = 1_048_575
# Only explicit simulation-name fields are valid scenario selectors.  In
# particular, OPTIONSET is a measure-set name and SIMULATION/SCENARIO/PLAN/RUN
# columns are not used because real ESP extracts use those words for other
# purposes too.
SIMULATION_COLUMNS = ("SIMULATION_NAME", "SIMULATIONNAME")


def output_folder(esp_folder: str | Path, create: bool = True) -> Path:
    path = Path(esp_folder).expanduser() / APP_OUTPUT_DIRNAME
    if create: path.mkdir(parents=True, exist_ok=True)
    return path


def timestamped_output_path(esp_folder: str | Path, stem: str, suffix: str) -> Path:
    now = datetime.now(); stamp = now.strftime("%d%b_%H%M")
    folder = output_folder(esp_folder); candidate = folder / f"{stem}_{stamp}{suffix}"
    if not candidate.exists(): return candidate
    candidate = folder / f"{stem}_{stamp}_{now:%S}{suffix}"
    counter = 2
    while candidate.exists():
        candidate = folder / f"{stem}_{stamp}_{now:%S}_{counter}{suffix}"; counter += 1
    return candidate

ALIASES: dict[str, tuple[str, ...]] = {
    "DATE": ("DATE", "STARTDATE", "START_DATE", "NEEDDATE", "NEED_DATE", "DMD_DATE",
             "FCSTDATE", "FCST_DATE", "ORDERDATE", "ORDER_DATE", "DUEDATE", "DUE_DATE",
             "RECEIPTDATE", "RECEIPT_DATE", "AVAILDATE", "AVAIL_DATE", "SUPPLYAVAILDATE",
             "SCHEDARRIVDATE", "SCHEDSHIPDATE", "DMDSTARTDATE", "FCSTSTARTDATE", "BUCKETDATE",
             "PERIOD", "MONTH", "WEEK"),
    "BU": ("BU", "BUSINESSUNIT", "BUSINESS_UNIT", "U_BU", "U_BUSINESS_UNIT", "PRODUCTLVL_1",
           "U_PRODUCTLVL_1", "BU_FCST"),
    "PLANT": ("PLANT", "LOCATION", "LOC", "DFULOC", "DEST", "DESTINATION", "SITE", "U_PLANT", "U_LOC"),
    "RESOURCE": ("RESOURCE", "RES", "RESNAME", "RESOURCE_NAME", "U_RESOURCE"),
    "PRODUCT_FORM": ("PRODUCT_FORM", "PRODUCTFORM", "FORM", "U_PRODUCT_FORM", "U_PRODUCTFORM"),
    "STREAM": ("STREAM", "U_STREAM", "PRODUCTION_STREAM", "MFG_STREAM"),
    "SUB_CLASS": ("SUB_CLASS", "SUBCLASS", "U_SUB_CLASS", "PRODUCT_SUB_CLASS"),
    "PRODUCTLVL_1": ("U_PRODUCTLVL_1", "PRODUCTLVL_1", "PRODUCT_LEVEL_1"),
    "ITEM": ("ITEM", "SKU", "ITEMID", "ITEM_ID", "MATERIAL", "PRODUCT", "SKUID", "TOITEM", "SKUITEM", "DESTITEM", "OUTPUTITEM"),
    "DFU": ("DFU", "DFUNAME", "DFU_NAME", "DMDUNIT", "DEMANDUNIT", "DEMAND_FORECAST_UNIT"),
    "DMDGROUP": ("DMDGROUP", "DMD_GROUP", "DEMANDGROUP", "DEMAND_GROUP"),
    "CUSTOMER": ("CUSTOMER", "CUST", "CUSTOMER_ID", "CUSTID"),
    "ORDER_ID": ("ORDER_ID", "ORDERID", "ORDERNUMBER", "ORDER_NUMBER", "ORDERNO"),
    "PRIORITY": ("PRIORITY", "ORDER_PRIORITY", "ORDERPRIORITY", "PRIORITYCODE"),
    "QTY": ("QTY", "QUANTITY", "VALUE", "PLANQTY", "PLAN_QTY", "FCSTQTY", "FCST_QTY", "TOTFCST",
            "DMDQTY", "DMD_QTY", "ORDERQTY", "ORDER_QTY", "SCHEDQTY", "SCHED_QTY",
            "PRODQTY", "PROD_QTY", "TOTALQTY", "TOTAL_QTY", "SUPPLYPEGQTY", "OUTPUT_QTY", "SBU_FCST"),
    "CAPACITY": ("CAPACITY", "AVAIL_CAPACITY", "AVAILABLE_CAPACITY", "CAPQTY", "CAP_QTY"),
    "LOAD": ("LOAD", "RESOURCE_LOAD", "REQUIRED_CAPACITY", "REQ_CAPACITY", "LOADQTY", "LOAD_QTY"),
}

MEASURE_ALIASES: dict[str, tuple[str, ...]] = {
    "Opening Inventory": ("OPENING_INVENTORY", "OPENINGINV", "OPEN_INV", "BEGINNING_INVENTORY", "BEGIN_INV"),
    "Purchase Receipt": ("PURCHASE_RECEIPT", "PURCH_RECEIPT", "PURCHASE", "RECEIPT", "INBOUND_RECEIPT"),
    "Production": ("PRODUCTION", "PROD_QTY", "PRODUCTION_QTY", "SCHED_QTY", "SCHEDULED_QTY", "SUPPLY"),
    "Dispatch": ("DISPATCH", "DISPATCH_QTY", "SHIPMENT", "SHIP_QTY", "CUSTOMER_ORDER", "CUSTORDER"),
    "In-Transit": ("IN_TRANSIT", "INTRANSIT", "TRANSIT_QTY", "INTRANSIT_QTY"),
    "Closing Inventory": ("CLOSING_INVENTORY", "CLOSINGINV", "CLOSE_INV", "ENDING_INVENTORY", "END_INV"),
}

FORECAST_MEASURES: dict[str, tuple[str, ...]] = {
    "BU_FCST": ("SBU_FCST", "BU_FCST", "BU_FORECAST", "FCST_BU"),
    "DFUTOSKU_FCST": ("TOTFCST", "DFUTOSKU_FCST", "SKU_FCST", "SKU_FORECAST"),
    "NETTED_FCST": ("NETTED_FCST", "NET_FCST", "NET_FORECAST"),
    "CUSTORDER": ("CUSTORDER", "CUSTOMER_ORDER", "ORDER_QTY", "CUST_ORDER"),
    "TOTAL_FCST": ("TOTAL_FCST", "TOTAL_FORECAST", "TOTAL_DEMAND"),
    "SCHED_QTY": ("SCHED_QTY", "SCHEDULED_QTY", "SUPPLY_QTY", "PLAN_QTY"),
}

FILE_CLASS_KEYWORDS = {
    "bu_head_forecast": ("UDT_BU_HEAD_FCST", "BU_HEAD_FCST"),
    "inventory": ("INVENTORY", "INV_BAL", "INVBAL", "STOCK"),
    "purchase": ("PURCHASE", "PURCH", "RECEIPT", "PO_"),
    "production": ("PRODUCTION", "PROD", "SCHEDULE", "SCHED"),
    "dispatch": ("DISPATCH", "SHIPMENT", "SHIP", "CUSTORDER", "CUSTOMER_ORDER"),
    "intransit": ("INTRANSIT", "IN_TRANSIT", "TRANSIT"),
    "resource": ("RESOURCE", "RESLOAD", "CAPACITY", "BOTTLENECK"),
    "demand": ("DEMAND", "DMD", "FCST", "FORECAST", "INDDMDVIEW"),
    "order": ("ORDER", "PRIORITY", "FULFIL"),
    "mapping": ("DFUTOSKU_FE", "DFUTOSKU_MAP", "MAPPING"),
}


def normalise_name(value: Any) -> str:
    text = str(value).strip().upper()
    text = re.sub(r"[^A-Z0-9]+", "_", text)
    return re.sub(r"_+", "_", text).strip("_") or "UNNAMED"


def unique_columns(columns: Iterable[Any]) -> list[str]:
    seen: dict[str, int] = {}
    result: list[str] = []
    for raw in columns:
        base = normalise_name(raw)
        seen[base] = seen.get(base, 0) + 1
        result.append(base if seen[base] == 1 else f"{base}_{seen[base]}")
    return result


def first_column(df: pd.DataFrame, aliases: Sequence[str], contains: bool = False) -> str | None:
    cols = list(df.columns)
    for candidate in aliases:
        c = normalise_name(candidate)
        if c in cols:
            return c
    if contains:
        for candidate in aliases:
            c = normalise_name(candidate)
            for col in cols:
                if c in col:
                    return col
    return None


def canonical_column(df: pd.DataFrame, field: str) -> str | None:
    return first_column(df, ALIASES[field], contains=False)


def clean_text(series: pd.Series, default: str = "(Blank)") -> pd.Series:
    out = series.astype("string").str.strip()
    return out.mask(out.isna() | out.eq("") | out.str.upper().isin(["NAN", "NONE", "NAT"]), default)


def parse_dates(series: pd.Series) -> pd.Series:
    """Return strictly ``datetime64[ns]`` dates bounded to 1900-2200.

    Pandas 3 can parse an out-of-range value to a coarse datetime unit and then
    fail while assigning it to a nanosecond Series.  Values such as ESP's
    999999/9999-12-31 placeholders are therefore rejected *before* conversion.
    """
    result = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")
    raw = series.astype("string").str.strip()
    num = pd.to_numeric(raw, errors="coerce").astype("float64")

    serial = num.between(1, 109_575).fillna(False)  # Excel serials through 2200-01-01.
    if serial.any():
        serial_dates = pd.Timestamp("1899-12-30") + pd.to_timedelta(num.loc[serial], unit="D")
        serial_dates = serial_dates.where(serial_dates.dt.year.between(1900, 2200))
        result.loc[serial] = serial_dates.astype("datetime64[ns]")

    yyyymm = num.between(190001, 220012) & ~serial
    if yyyymm.any():
        parsed = pd.to_datetime(num.loc[yyyymm].astype("Int64").astype(str), format="%Y%m", errors="coerce")
        result.loc[yyyymm] = parsed.astype("datetime64[ns]")

    other = ~serial & ~yyyymm
    if other.any():
        text = raw.loc[other]
        # Reject an explicit 4-digit year outside the supported range before
        # pandas attempts high-resolution conversion.
        years = pd.to_numeric(text.str.extract(r"(?<!\d)(\d{4})(?!\d)", expand=False), errors="coerce")
        safe = years.isna() | years.between(1900, 2200)
        parse_text = text.str.replace(r"(?<=\d)_(?=\d)", "-", regex=True)
        parsed = pd.to_datetime(parse_text.where(safe), errors="coerce", format="mixed", dayfirst=True)
        failed = safe & parsed.isna()
        if failed.any():
            # Normalised wide headers such as 22_OCT_26_12_00_AM.
            header_text = text.loc[failed].str.replace("_", " ", regex=False)
            parsed.loc[failed] = pd.to_datetime(header_text, errors="coerce", format="mixed", dayfirst=True)
        parsed = parsed.where(parsed.dt.year.between(1900, 2200))
        result.loc[other] = parsed.astype("datetime64[ns]")
    return result.astype("datetime64[ns]")


def numeric(series: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce").fillna(0.0)
    cleaned = (series.astype("string").str.replace(",", "", regex=False)
               .str.replace("(", "-", regex=False).str.replace(")", "", regex=False)
               .str.replace(r"[^0-9eE+\-.]", "", regex=True))
    return pd.to_numeric(cleaned, errors="coerce").fillna(0.0)


def quantity_mt(series: pd.Series, source_column: str = "") -> pd.Series:
    values = numeric(series)
    c = normalise_name(source_column)
    return values if re.search(r"(^|_)MT($|_)", c) else values / 1000.0


def week_start(dates: pd.Series) -> pd.Series:
    d = pd.to_datetime(dates, errors="coerce")
    # ESP uses 1/8/15/22 buckets. Preserve those exact dates and fold arbitrary
    # transaction dates (for example customer SHIPDATE) into the containing ESP
    # bucket instead of creating a misleading daily axis.
    day = d.dt.day
    bucket_day = np.select([day.ge(22), day.ge(15), day.ge(8), day.ge(1)], [22, 15, 8, 1], default=1)
    out = pd.to_datetime(pd.DataFrame({"year": d.dt.year, "month": d.dt.month, "day": bucket_day}, index=d.index), errors="coerce")
    return out.astype("datetime64[ns]")


def month_start(dates: pd.Series) -> pd.Series:
    d = pd.to_datetime(dates, errors="coerce")
    return d.dt.to_period("M").dt.to_timestamp()


def classify_file(name: str) -> str:
    n = normalise_name(Path(name).stem)
    best = "other"
    best_score = 0
    for kind, keywords in FILE_CLASS_KEYWORDS.items():
        score = sum(len(k) for k in keywords if k in n)
        if score > best_score:
            best, best_score = kind, score
    return best


def _read_csv(path: Path) -> pd.DataFrame:
    last_error: Exception | None = None
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin1"):
        try:
            # The C parser is dramatically faster for the large ESP extracts.
            frame = pd.read_csv(path, encoding=encoding, low_memory=False)
            if len(frame.columns) > 1 or not any(delimiter in str(frame.columns[0]) for delimiter in ("\t", ";", "|")):
                return frame
            # A one-column result whose header still contains a delimiter was
            # not comma-separated; continue to the sniffing fallback below.
            last_error = ValueError("C parser did not detect the delimiter")
            break
        except Exception as exc:  # pragma: no cover - depends on external encodings
            last_error = exc
    # Delimiter sniffing is a last resort for semicolon/tab/custom exports.
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin1"):
        try:
            return pd.read_csv(path, encoding=encoding, sep=None, engine="python")
        except Exception as exc:  # pragma: no cover
            last_error = exc
    raise ValueError(f"Could not read {path.name}: {last_error}")


def _apply_simulation(df: pd.DataFrame, simulation: str | None) -> tuple[pd.DataFrame, str]:
    if not simulation:
        return df, "No simulation filter"
    target = simulation.strip().casefold()
    candidates = [c for c in SIMULATION_COLUMNS if c in df.columns]
    if not candidates:
        return df, "No simulation column; file retained"
    masks: list[pd.Series] = []
    for col in candidates:
        masks.append(clean_text(df[col], "").str.casefold().eq(target))
    mask = pd.concat(masks, axis=1).any(axis=1)
    if mask.any():
        return df.loc[mask].copy(), f"Filtered on {', '.join(candidates)}"
    return df.iloc[0:0].copy(), f"Simulation {simulation!r} not present"


@dataclass
class SourceTable:
    name: str
    path: Path
    kind: str
    data: pd.DataFrame
    original_columns: list[str]
    simulation_note: str
    available_simulations: tuple[str, ...] = ()
    rows_before_filter: int = 0
    rows_after_filter: int = 0
    load_error: str = ""


@dataclass
class TwinModel:
    folder: Path
    simulation: str | None
    sources: dict[str, SourceTable] = field(default_factory=dict)
    tables: dict[str, pd.DataFrame] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)
    generated_at: datetime = field(default_factory=datetime.now)

    @property
    def source_rows(self) -> int:
        return int(sum(len(s.data) for s in self.sources.values()))

    def kpis(self) -> dict[str, float | int]:
        trend = self.tables.get("weekly_trend", pd.DataFrame())
        lost = self.tables.get("dfu_lost_mapping", pd.DataFrame())
        capacity = self.tables.get("capacity", self.tables.get("bottlenecks", pd.DataFrame()))
        service = self.tables.get("service_demand", pd.DataFrame())
        orders = self.tables.get("order_status", pd.DataFrame())
        margin = self.tables.get("value_at_risk", pd.DataFrame())
        def total_measure(measure: str) -> float:
            if trend.empty: return 0.0
            return float(trend.loc[trend["MEASURE"].eq(measure), "QTY_MT"].sum())
        closing = trend[trend.get("MEASURE", pd.Series(index=trend.index, dtype=str)).eq("Closing Inventory")].sort_values("WEEK") if not trend.empty else trend
        latest_closing = (closing.groupby(["PLANT", "ITEM"], dropna=False, observed=True).last()["QTY_MT"].sum() if not closing.empty else 0.0)
        demand = float(service.get("DEMAND_MT", pd.Series(dtype=float)).sum())
        allocated = float(service.get("ALLOCATED_MT", pd.Series(dtype=float)).sum())
        horizon = orders[orders.get("ELIGIBLE_IN_HORIZON", pd.Series(True, index=orders.index)).fillna(False)] if not orders.empty else orders
        otif_weight = float(horizon.get("OTIF_WEIGHT_MT", pd.Series(dtype=float)).sum())
        order_weight = float(horizon.get("ORDER_MT", pd.Series(dtype=float)).sum())
        fg = trend[trend.get("SOURCE_FILE", pd.Series(dtype=str)).astype(str).str.contains("FG_Production_Plan", case=False, na=False)] if not trend.empty else trend
        return {
            "Source files": len(self.sources),
            "Source rows": self.source_rows,
            "Production MT": total_measure("Production"),
            "FG plan MT": float(fg.loc[fg.get("MEASURE", pd.Series(index=fg.index, dtype=str)).eq("Production"), "QTY_MT"].sum()) if not fg.empty else 0.0,
            "Dispatch MT": total_measure("Dispatch"),
            "Closing inventory MT": float(latest_closing),
            "Order fill %": allocated / demand * 100 if demand else np.nan,
            "Unallocated MT": float(service.get("UNALLOCATED_MT", pd.Series(dtype=float)).sum()),
            "Forecast lost MT": float(lost.get("LOST_MT", pd.Series(dtype=float)).sum()),
            "DFU mapping issues": int(len(lost)),
            "Bottleneck resources": int(capacity.get("BOTTLENECK", pd.Series(dtype=bool)).fillna(False).sum()),
            "Resource periods": int(len(capacity)),
            "OTIF in horizon %": otif_weight / order_weight * 100 if order_weight else np.nan,
            "Negative-margin scheduled MT": float(margin.get("NEGATIVE_MARGIN_SCHEDULED_MT", pd.Series(dtype=float)).sum()),
        }


def load_sources(folder: str | Path, simulation: str | None = None) -> tuple[dict[str, SourceTable], list[str]]:
    root = Path(folder).expanduser()
    if not root.exists() or not root.is_dir():
        raise FileNotFoundError(f"Data folder does not exist: {root}")
    files = sorted([p for p in root.rglob("*")
                    if p.is_file() and p.suffix.lower() in {".csv", ".txt"}
                    and APP_OUTPUT_DIRNAME.casefold() not in {part.casefold() for part in p.relative_to(root).parts}])
    if not files:
        raise FileNotFoundError(f"No CSV/TXT files found under: {root}")
    sources: dict[str, SourceTable] = {}
    warnings: list[str] = []
    for path in files:
        try:
            df = _read_csv(path)
            original = [str(c) for c in df.columns]
            df.columns = unique_columns(df.columns)
            before = len(df)
            sim_values: set[str] = set()
            for sim_col in SIMULATION_COLUMNS:
                if sim_col in df.columns:
                    sim_values.update(clean_text(df[sim_col], "").loc[lambda x: x.ne("")].astype(str).unique().tolist())
            df, note = _apply_simulation(df, simulation)
            key = normalise_name(path.stem)
            if key in sources:
                i = 2
                while f"{key}_{i}" in sources: i += 1
                key = f"{key}_{i}"
            sources[key] = SourceTable(key, path, classify_file(path.name), df, original, note, tuple(sorted(sim_values)), before, len(df), "")
            if before and not len(df):
                warnings.append(f"{path.name}: 0 rows after simulation filter ({note}).")
        except Exception as exc:
            message = str(exc)
            warnings.append(f"{path.name}: {message}")
            key = normalise_name(path.stem)
            sources[key] = SourceTable(key, path, classify_file(path.name), pd.DataFrame(), [], "Load failed", (), 0, 0, message)
    if not sources:
        raise ValueError("Files were found, but none could be read.")
    return sources, warnings


def discover_simulations(folder: str | Path) -> list[str]:
    """Read only headers/simulation-name columns and return available scenarios."""
    root = Path(folder).expanduser()
    values: set[str] = set()
    if not root.exists():
        return []
    for path in root.rglob("*"):
        if not path.is_file() or path.suffix.lower() not in {".csv", ".txt"}:
            continue
        if APP_OUTPUT_DIRNAME.casefold() in {part.casefold() for part in path.relative_to(root).parts}:
            continue
        try:
            df = _read_csv(path)
            df.columns = unique_columns(df.columns)
            for col in SIMULATION_COLUMNS:
                if col in df:
                    values.update(clean_text(df[col], "").loc[lambda x: x.ne("")].astype(str).unique())
        except Exception:
            continue
    return sorted(values)


def _bu_from_item(item: pd.Series) -> pd.Series:
    text = clean_text(item, "").str.upper()
    out = pd.Series("(Blank)", index=item.index, dtype="string")
    out.loc[text.str.startswith("R_")] = "FRP"
    out.loc[text.str.startswith("F_")] = "Foils"
    out.loc[text.str.startswith("P_") | text.str.contains("HOT_METAL", regex=False)] = "PRIMARY"
    out.loc[text.str.match(r"^[EI](?:_|\d|[A-Z])", na=False)] = "E"
    return out


def _normalise_bu(series: pd.Series) -> pd.Series:
    text = clean_text(series, "").str.upper()
    mapped = text.replace({
        "R": "FRP", "FRP": "FRP", "F": "Foils", "FOIL": "Foils", "FOILS": "Foils",
        "P": "PRIMARY", "PRIMARY": "PRIMARY", "HOT METAL": "PRIMARY", "HOT_METAL": "PRIMARY",
        "E": "E", "I": "E",
    })
    return mapped.mask(mapped.isin(["", "(BLANK)", "NAN", "NONE", "NAT"]), "(Blank)")


def build_item_master(sources: Mapping[str, SourceTable]) -> pd.DataFrame:
    """Item → BU/product attributes, prioritising ESP_SKU and FG plan."""
    frames: list[pd.DataFrame] = []
    ranked = sorted(sources.values(), key=lambda s: (0 if s.name == "ESP_SKU" else 1 if "FG_PRODUCTION_PLAN" in s.name else 2))
    for source in ranked:
        df = source.data
        item_col = canonical_column(df, "ITEM")
        if not item_col or df.empty:
            continue
        part = pd.DataFrame({"ITEM": clean_text(df[item_col])})
        for field in ("BU", "PRODUCT_FORM", "STREAM", "SUB_CLASS", "PRODUCTLVL_1"):
            col = canonical_column(df, field)
            part[field] = clean_text(df[col]) if col else "(Blank)"
        part["BU"] = _normalise_bu(part["BU"])
        blank = part["BU"].eq("(Blank)")
        part.loc[blank, "BU"] = _normalise_bu(part.loc[blank, "PRODUCTLVL_1"])
        blank = part["BU"].eq("(Blank)")
        part.loc[blank, "BU"] = _bu_from_item(part.loc[blank, "ITEM"])
        container = clean_text(part["PRODUCTLVL_1"], "").str.upper().eq("E")
        part.loc[container, "SUB_CLASS"] = "Container"
        part["MASTER_SOURCE"] = source.path.name
        frames.append(part.drop_duplicates("ITEM"))
    if not frames:
        return pd.DataFrame(columns=["ITEM", "BU", "PRODUCT_FORM", "STREAM", "SUB_CLASS", "PRODUCTLVL_1", "MASTER_SOURCE"])
    combined = pd.concat(frames, ignore_index=True)
    # Stable ranking above means the first nonblank value is authoritative.
    rows = []
    for item, group in combined.groupby("ITEM", sort=False, dropna=False):
        row = {"ITEM": item}
        for col in ("BU", "PRODUCT_FORM", "STREAM", "SUB_CLASS", "PRODUCTLVL_1", "MASTER_SOURCE"):
            valid = group.loc[~group[col].isin(["(Blank)", ""]), col]
            row[col] = valid.iloc[0] if len(valid) else "(Blank)"
        rows.append(row)
    return pd.DataFrame(rows)


def enrich_with_item_master(frame: pd.DataFrame, item_master: pd.DataFrame) -> pd.DataFrame:
    if frame.empty or "ITEM" not in frame or item_master.empty:
        if "LOCATION_TYPE" not in frame: frame["LOCATION_TYPE"] = np.where(frame.get("PLANT", pd.Series(dtype=str)).astype(str).str.upper().str.startswith("V_"), "Demand zone", "Plant") if len(frame) else pd.Series(dtype="string")
        return frame
    attrs = [c for c in ["ITEM", "BU", "PRODUCT_FORM", "STREAM", "SUB_CLASS"] if c in item_master]
    out = frame.merge(item_master[attrs].drop_duplicates("ITEM"), on="ITEM", how="left", suffixes=("", "_MASTER"))
    for col in ("BU", "PRODUCT_FORM", "STREAM", "SUB_CLASS"):
        master = col + "_MASTER"
        if master in out:
            blank = out[col].isna() | out[col].astype(str).isin(["", "(Blank)", "nan", "None"])
            out.loc[blank, col] = out.loc[blank, master]
            out[col] = out[col].fillna("(Blank)")
            out = out.drop(columns=master)
    blank_bu = out["BU"].astype(str).eq("(Blank)")
    out.loc[blank_bu, "BU"] = _bu_from_item(out.loc[blank_bu, "ITEM"])
    out["LOCATION_TYPE"] = np.where(clean_text(out["PLANT"], "").str.upper().str.startswith("V_"), "Demand zone", "Plant")
    return out


def dimension_frame(df: pd.DataFrame) -> pd.DataFrame:
    out = pd.DataFrame(index=df.index)
    for field in ("BU", "PLANT", "RESOURCE", "PRODUCT_FORM", "STREAM", "ITEM", "DFU", "DMDGROUP"):
        col = canonical_column(df, field)
        out[field] = clean_text(df[col]) if col else "(Blank)"
    out["BU"] = _normalise_bu(out["BU"])
    sub = canonical_column(df, "SUB_CLASS")
    lvl1 = canonical_column(df, "PRODUCTLVL_1")
    out["SUB_CLASS"] = clean_text(df[sub]) if sub else "(Blank)"
    if lvl1:
        container = clean_text(df[lvl1], "").str.upper().eq("E")
        out.loc[container, "SUB_CLASS"] = "Container"
        blank = out["BU"].eq("(Blank)")
        out.loc[blank, "BU"] = _normalise_bu(df.loc[blank, lvl1])
    blank = out["BU"].eq("(Blank)")
    out.loc[blank, "BU"] = _bu_from_item(out.loc[blank, "ITEM"])
    out["LOCATION_TYPE"] = np.where(clean_text(out["PLANT"], "").str.upper().str.startswith("V_"), "Demand zone", "Plant")
    return out


def _measure_columns(df: pd.DataFrame) -> dict[str, str]:
    found: dict[str, str] = {}
    for label, aliases in MEASURE_ALIASES.items():
        col = first_column(df, aliases, contains=False)
        if col: found[label] = col
    return found


def _default_measure(kind: str, filename: str) -> str | None:
    n = normalise_name(filename)
    if kind == "inventory":
        if "OPEN" in n or "BEGIN" in n: return "Opening Inventory"
        if "CLOS" in n or "END" in n: return "Closing Inventory"
    return {"purchase": "Purchase Receipt", "production": "Production", "dispatch": "Dispatch", "intransit": "In-Transit"}.get(kind)


def _find_value_column(df: pd.DataFrame, preferred: Sequence[str], excluded: Sequence[str] = ()) -> str | None:
    col = first_column(df, preferred)
    if col:
        return col
    excluded_set = set(excluded)
    scores: list[tuple[float, str]] = []
    for candidate in df.columns:
        if candidate in excluded_set:
            continue
        converted = pd.to_numeric(df[candidate], errors="coerce")
        score = float(converted.notna().mean())
        if score >= .5:
            scores.append((score, candidate))
    return max(scores, default=(0.0, None))[1]


def _make_fact(source: SourceTable, label: str, qty_aliases: Sequence[str], date_aliases: Sequence[str],
               plant_aliases: Sequence[str] | None = None, rows: pd.Series | None = None) -> pd.DataFrame:
    df = source.data if rows is None else source.data.loc[rows]
    if df.empty:
        return pd.DataFrame()
    date_col = first_column(df, date_aliases)
    qty_col = _find_value_column(df, qty_aliases, excluded=[date_col] if date_col else [])
    if not date_col or not qty_col:
        return pd.DataFrame()
    part = dimension_frame(df)
    if plant_aliases:
        plant_col = first_column(df, plant_aliases)
        if plant_col:
            part["PLANT"] = clean_text(df[plant_col])
            part["LOCATION_TYPE"] = np.where(part["PLANT"].str.upper().str.startswith("V_"), "Demand zone", "Plant")
    part["DATE"] = parse_dates(df[date_col])
    part["WEEK"] = week_start(part["DATE"])
    part["MONTH"] = month_start(part["DATE"])
    part["MEASURE"] = label
    part["QTY_MT"] = quantity_mt(df[qty_col], qty_col)
    part["SOURCE_FILE"] = source.path.name
    part["MATCHED_DATE_COLUMN"] = date_col
    part["MATCHED_QTY_COLUMN"] = qty_col
    return part.loc[part["DATE"].notna()]


def _resource_output_fact(source: SourceTable) -> pd.DataFrame:
    df = source.data
    if df.empty:
        return pd.DataFrame()
    measure_col = None
    for col in df.columns:
        values = clean_text(df[col], "").str.upper().str.replace(r"\s+", " ", regex=True)
        if values.isin(["% RESOURCE UTIL", "INPUT QTY", "OUTPUT QTY", "CO_PRODUCT_QTY", "CO PRODUCT QTY"]).any():
            measure_col = col
            break
    if not measure_col:
        return pd.DataFrame()
    values = clean_text(df[measure_col], "").str.upper().str.replace(r"\s+", " ", regex=True)
    output_rows = values.isin(["OUTPUT QTY", "CO_PRODUCT_QTY", "CO PRODUCT QTY"])
    date_col = first_column(df, ("DATE", "STARTDATE", "WEEK", "MONTH", "BUCKETDATE"))
    qty_col = _find_value_column(df.loc[output_rows], ("QTY", "QUANTITY", "VALUE", "OUTPUT_QTY"), excluded=[measure_col, date_col or ""])
    if date_col and qty_col:
        return _make_fact(source, "Production", (qty_col,), (date_col,), rows=output_rows)

    # Some UDT exports are wide: dimension columns followed by date columns.
    date_columns: list[str] = []
    parsed_headers: dict[str, pd.Timestamp] = {}
    for col in df.columns:
        parsed = parse_dates(pd.Series([col])).iloc[0]
        if pd.notna(parsed):
            date_columns.append(col); parsed_headers[col] = parsed
    if not date_columns:
        return pd.DataFrame()
    base = df.loc[output_rows].copy()
    id_cols = [c for c in base.columns if c not in date_columns]
    melted = base.melt(id_vars=id_cols, value_vars=date_columns, var_name="FACT_DATE", value_name="FACT_QTY")
    melted["FACT_DATE"] = melted["FACT_DATE"].map(parsed_headers)
    temp = SourceTable(source.name, source.path, source.kind, melted, source.original_columns, source.simulation_note, source.available_simulations)
    return _make_fact(temp, "Production", ("FACT_QTY",), ("FACT_DATE",))


def build_weekly_trend(sources: Mapping[str, SourceTable]) -> pd.DataFrame:
    """Build the material-flow fact from the named ESP operational exports."""
    frames: list[pd.DataFrame] = []
    has_fg_plan = any("FG_PRODUCTION_PLAN" in s.name and not s.data.empty for s in sources.values())
    for source in sources.values():
        n, df = source.name, source.data
        if df.empty:
            continue

        # Keep compatibility with a purpose-built balance extract when present.
        measures = _measure_columns(df)
        if len(measures) >= 3:
            for label, qty_col in measures.items():
                part = _make_fact(source, label, (qty_col,), ALIASES["DATE"])
                if not part.empty: frames.append(part)
            continue

        if "FG_PRODUCTION_PLAN" in n:
            part = _make_fact(source, "Production", ("SUPPLYPEGQTY", "QTY"), ("SUPPLYAVAILDATE", "STARTDATE", "NEEDDATE"))
            if not part.empty: frames.append(part)
        elif "ESP_PLANORDER" in n and not has_fg_plan:
            part = _make_fact(source, "Production", ("QTY", "PLANQTY"), ("STARTDATE", "NEEDDATE"))
            if not part.empty: frames.append(part)
        elif "UDT_RES_UTIL_AND_ALLOC" in n and not has_fg_plan:
            part = _resource_output_fact(source)
            if not part.empty: frames.append(part)

        if n == "ESP_INVENTORY" or ("ESP_INVENTORY" in n and "PARAM" not in n and "BALANCE" not in n):
            close = _make_fact(source, "Closing Inventory", ("QTY", "PROJECTED_ON_HAND", "ONHANDQTY"), ("AVAILDATE", "DATE"))
            if not close.empty:
                close = close.sort_values(["ITEM", "PLANT", "DATE"])
                opening = close.copy(); opening["MEASURE"] = "Opening Inventory"
                opening["QTY_MT"] = close.groupby(["ITEM", "PLANT"], dropna=False, observed=True)["QTY_MT"].shift(1).fillna(close["QTY_MT"])
                frames.extend([opening, close])
        if "ESP_PURCHORDER" in n:
            part = _make_fact(source, "Purchase Receipt", ("QTY", "ORDERQTY", "REMAININGQTY"), ("DUEDATE", "NEEDDATE", "DATE"), ("LOC", "DEST", "LOCATION"))
            if not part.empty: frames.append(part)
        if "ESP_PLANARRIV" in n:
            transit = _make_fact(source, "In-Transit", ("QTY", "PLANQTY"), ("SCHEDARRIVDATE", "NEEDDATE"), ("DEST", "DESTINATION", "LOC"))
            dispatch = _make_fact(source, "Dispatch", ("QTY", "PLANQTY"), ("SCHEDSHIPDATE", "STARTDATE"), ("SOURCE", "ORIGIN"))
            if not transit.empty: frames.append(transit)
            if not dispatch.empty: frames.append(dispatch)
        if "ESP_VEHICLELOADLINE" in n or "VEHICLELOADLINE" in n:
            transit = _make_fact(source, "In-Transit", ("QTY", "SHIPQTY", "LOADQTY"), ("SCHEDARRIVDATE", "ARRIVDATE", "DUEDATE"), ("DEST", "DESTINATION", "LOC"))
            dispatch = _make_fact(source, "Dispatch", ("QTY", "SHIPQTY", "LOADQTY"), ("SCHEDSHIPDATE", "SHIPDATE", "STARTDATE"), ("SOURCE", "ORIGIN"))
            if not transit.empty: frames.append(transit)
            if not dispatch.empty: frames.append(dispatch)
        if "ORDER_FULFILLMENT" in n:
            part = _make_fact(source, "Dispatch", ("SHIPPEDQTY", "FULFILLEDQTY", "ALLOCATEDQTY"), ("SHIPDATE", "DUEDATE", "NEEDDATE"))
            if not part.empty: frames.append(part)

    columns = ["WEEK", "MONTH", "BU", "PLANT", "LOCATION_TYPE", "RESOURCE", "PRODUCT_FORM", "STREAM",
               "SUB_CLASS", "ITEM", "DFU", "DMDGROUP", "MEASURE", "QTY_MT", "SOURCE_FILE",
               "MATCHED_DATE_COLUMN", "MATCHED_QTY_COLUMN"]
    if not frames:
        return pd.DataFrame(columns=columns)
    all_rows = enrich_with_item_master(pd.concat(frames, ignore_index=True), build_item_master(sources))
    dims = ["WEEK", "MONTH", "BU", "PLANT", "LOCATION_TYPE", "RESOURCE", "PRODUCT_FORM", "STREAM", "SUB_CLASS", "ITEM", "DFU", "DMDGROUP", "MEASURE"]
    all_rows = all_rows.sort_values("DATE")
    flow = all_rows[~all_rows["MEASURE"].isin(["Opening Inventory", "Closing Inventory"])]
    stock = all_rows[all_rows["MEASURE"].isin(["Opening Inventory", "Closing Inventory"])]
    result = []
    if not flow.empty:
        result.append(flow.groupby(dims, dropna=False, observed=True, as_index=False).agg(
            QTY_MT=("QTY_MT", "sum"), SOURCE_FILE=("SOURCE_FILE", lambda x: ", ".join(sorted(set(map(str, x))))),
            MATCHED_DATE_COLUMN=("MATCHED_DATE_COLUMN", "first"), MATCHED_QTY_COLUMN=("MATCHED_QTY_COLUMN", "first")))
    if not stock.empty:
        opening = stock[stock["MEASURE"].eq("Opening Inventory")].groupby(dims, dropna=False, observed=True, as_index=False).first()
        closing = stock[stock["MEASURE"].eq("Closing Inventory")].groupby(dims, dropna=False, observed=True, as_index=False).last()
        result.extend([opening, closing])
    out = pd.concat(result, ignore_index=True) if result else pd.DataFrame(columns=columns)
    # The weekly axis must begin at the ESP run's inventory plan start, not at
    # historical order/ship dates.  ESP_INVENTORY AVAILDATE is authoritative.
    inv_start = _inventory_plan_start(sources)
    plan_end = _resource_plan_end(sources)
    if inv_start is not None and not out.empty:
        out = out[out["WEEK"].ge(inv_start)].copy()
    if plan_end is not None and not out.empty:
        out = out[out["WEEK"].le(plan_end)].copy()
    return out[columns].sort_values(["WEEK", "MEASURE", "PLANT", "ITEM"], kind="stable").reset_index(drop=True)


def build_resource_production(sources: Mapping[str, SourceTable], item_master: pd.DataFrame) -> pd.DataFrame:
    """Authoritative resource production from ESP_RESLOADDETAIL Output Qty.

    Each item-location is assigned to the resource with the largest horizon
    Output Qty.  This removes duplicate route rows and matches the inventory
    balance resource assignment.  PLANORDER/PRODUCTIONSTEP is used only when
    RESLOADDETAIL is missing.
    """
    source = _latest_matching_source(sources, ("ESP_RESLOADDETAIL", "RESLOADDETAIL"))
    columns = ["WEEK", "MONTH", "BU", "PLANT", "RESOURCE", "PRODUCT_FORM", "STREAM", "SUB_CLASS", "ITEM", "MEASURE", "QTY_MT"]
    if source is None or source.data.empty:
        return pd.DataFrame(columns=columns)
    df = source.data
    res = canonical_column(df, "RESOURCE"); loc = canonical_column(df, "PLANT"); item = canonical_column(df, "ITEM")
    form = canonical_column(df, "PRODUCT_FORM")
    when = first_column(df, ("WHENLOADED", "WHEN_LOADED", "STARTDATE", "DATE"))
    outqty = first_column(df, ("OUTPUT_QTY", "OUTPUT QTY", "OUTPUTQTY"))
    if not all((res, loc, item, when, outqty)):
        return pd.DataFrame(columns=columns)
    out = pd.DataFrame({
        "RESOURCE": clean_text(df[res]), "PLANT": clean_text(df[loc]), "ITEM": clean_text(df[item]),
        "PRODUCT_FORM": clean_text(df[form]) if form else "(Blank)", "DATE": parse_dates(df[when]),
        "QTY_MT": quantity_mt(df[outqty], outqty),
    })
    out = out[~out["RESOURCE"].str.upper().str.contains("DUMMY", regex=False)]
    totals = out.groupby(["ITEM", "PLANT", "RESOURCE"], observed=True, as_index=False)["QTY_MT"].sum()
    winners = totals.loc[totals.groupby(["ITEM", "PLANT"])["QTY_MT"].idxmax(), ["ITEM", "PLANT", "RESOURCE"]]
    out = out.merge(winners.assign(_WIN=True), on=["ITEM", "PLANT", "RESOURCE"], how="inner")
    out["WEEK"] = week_start(out["DATE"]); out["MONTH"] = month_start(out["DATE"])
    out["BU"] = _bu_from_item(out["ITEM"]); out["STREAM"] = "(Blank)"; out["SUB_CLASS"] = "(Blank)"; out["MEASURE"] = "Production"
    if not item_master.empty:
        attrs = item_master[[c for c in ["ITEM", "BU", "PRODUCT_FORM", "STREAM", "SUB_CLASS"] if c in item_master]].drop_duplicates("ITEM")
        out = out.merge(attrs, on="ITEM", how="left", suffixes=("", "_MASTER"))
        for c in ["BU", "PRODUCT_FORM", "STREAM", "SUB_CLASS"]:
            mc = c + "_MASTER"
            if mc in out:
                blank = out[c].astype(str).isin(["", "(Blank)", "nan", "None"])
                out.loc[blank, c] = out.loc[blank, mc]
                out = out.drop(columns=mc)
    start, end = _inventory_plan_start(sources), _resource_plan_end(sources)
    if start is not None: out = out[out["WEEK"].ge(start)]
    if end is not None: out = out[out["WEEK"].le(end)]
    return (out.groupby(columns[:-1], dropna=False, observed=True, as_index=False)["QTY_MT"].sum()
            if not out.empty else pd.DataFrame(columns=columns))


def aggregate_trend(trend: pd.DataFrame, period: str, dimensions: Sequence[str], measures: Sequence[str] | None = None) -> pd.DataFrame:
    if trend.empty: return pd.DataFrame(columns=[period, *dimensions, "MEASURE", "QTY_MT"])
    work = trend
    if measures: work = work[work["MEASURE"].isin(measures)]
    keys = [period, *[d for d in dimensions if d in work.columns], "MEASURE"]
    stocks = work[work["MEASURE"].isin(["Opening Inventory", "Closing Inventory"])]
    flows = work[~work["MEASURE"].isin(["Opening Inventory", "Closing Inventory"])]
    parts = []
    if not flows.empty:
        parts.append(flows.groupby(keys, dropna=False, observed=True, as_index=False)["QTY_MT"].sum())
    if not stocks.empty:
        opener = stocks[stocks["MEASURE"].eq("Opening Inventory")].sort_values("WEEK").groupby(keys, dropna=False, observed=True, as_index=False).first()
        closer = stocks[stocks["MEASURE"].eq("Closing Inventory")].sort_values("WEEK").groupby(keys, dropna=False, observed=True, as_index=False).last()
        parts.extend([opener[keys + ["QTY_MT"]], closer[keys + ["QTY_MT"]]])
    return (pd.concat(parts, ignore_index=True).sort_values(keys, kind="stable").reset_index(drop=True)
            if parts else pd.DataFrame(columns=[*keys, "QTY_MT"]))


def _find_sources(sources: Mapping[str, SourceTable], required: Sequence[str]) -> list[SourceTable]:
    req = [normalise_name(x) for x in required]
    scored: list[tuple[int, SourceTable]] = []
    for source in sources.values():
        n = source.name
        score = sum(len(token) for token in req if token in n)
        if score: scored.append((score, source))
    return [s for _, s in sorted(scored, key=lambda x: x[0], reverse=True)]


def _latest_matching_source(sources: Mapping[str, SourceTable], tokens: Sequence[str]) -> SourceTable | None:
    candidates = [s for s in sources.values() if any(normalise_name(t) in s.name for t in tokens)]
    if not candidates: return None
    def version_key(source: SourceTable) -> tuple[str, int, str]:
        stamps = re.findall(r"(?<!\d)(\d{12,14})(?!\d)", source.path.stem)
        stamp = max(stamps) if stamps else ""
        mtime = source.path.stat().st_mtime_ns if source.path.exists() else 0
        return stamp, mtime, source.path.name
    return max(candidates, key=version_key)


def _dmdgroup_key(series: pd.Series) -> pd.Series:
    """Strip ESP stream suffixes: 20167481_STREAM → 20167481."""
    return clean_text(series).str.split("_", n=1).str[0].str.strip()


def _fact_from_source(source: SourceTable, value_label: str = "QTY_MT") -> pd.DataFrame:
    df = source.data
    if df.empty: return pd.DataFrame()
    out = dimension_frame(df)
    date_col = canonical_column(df, "DATE")
    out["DATE"] = parse_dates(df[date_col]) if date_col else pd.NaT
    out["WEEK"] = week_start(out["DATE"])
    out["MONTH"] = month_start(out["DATE"])
    qty_col = canonical_column(df, "QTY")
    out[value_label] = quantity_mt(df[qty_col], qty_col) if qty_col else 0.0
    for field in ("CUSTOMER", "ORDER_ID", "PRIORITY"):
        col = canonical_column(df, field)
        out[field] = clean_text(df[col]) if col else "(Blank)"
    out["SOURCE_FILE"] = source.path.name
    return out


def build_bottlenecks(sources: Mapping[str, SourceTable], trend: pd.DataFrame) -> pd.DataFrame:
    frames: list[pd.DataFrame] = []
    for source in sources.values():
        df = source.data
        cap = canonical_column(df, "CAPACITY")
        load = canonical_column(df, "LOAD")
        if not cap and not load: continue
        dims = dimension_frame(df)
        date_col = canonical_column(df, "DATE")
        dims["WEEK"] = week_start(parse_dates(df[date_col])) if date_col else pd.NaT
        dims["CAPACITY_MT"] = quantity_mt(df[cap], cap) if cap else np.nan
        dims["LOAD_MT"] = quantity_mt(df[load], load) if load else np.nan
        frames.append(dims[["WEEK", "BU", "PLANT", "RESOURCE", "STREAM", "CAPACITY_MT", "LOAD_MT"]])
    if frames:
        out = pd.concat(frames, ignore_index=True)
        keys = ["WEEK", "BU", "PLANT", "RESOURCE", "STREAM"]
        out = out.groupby(keys, dropna=False, observed=True, as_index=False).sum(min_count=1)
    else:
        prod = aggregate_trend(trend, "WEEK", ["BU", "PLANT", "RESOURCE", "STREAM"], ["Production"])
        if prod.empty:
            return pd.DataFrame(columns=["WEEK", "BU", "PLANT", "RESOURCE", "STREAM", "LOAD_MT", "CAPACITY_MT", "UTILIZATION_PCT", "BOTTLENECK"])
        out = prod.rename(columns={"QTY_MT": "LOAD_MT"}).drop(columns="MEASURE")
        out["CAPACITY_MT"] = np.nan
    out["UTILIZATION_PCT"] = np.where(out["CAPACITY_MT"].gt(0), out["LOAD_MT"] / out["CAPACITY_MT"] * 100, np.nan)
    out["BOTTLENECK"] = out["UTILIZATION_PCT"].ge(95)
    return out.sort_values(["BOTTLENECK", "UTILIZATION_PCT", "LOAD_MT"], ascending=[False, False, False], na_position="last").reset_index(drop=True)


def _table_by_exact_tokens(sources: Mapping[str, SourceTable], tokens: Sequence[str]) -> SourceTable | None:
    candidates = _find_sources(sources, tokens)
    return candidates[0] if candidates else None


def _id_series(df: pd.DataFrame, field: str, default: str = "(Blank)") -> pd.Series:
    col = canonical_column(df, field)
    return clean_text(df[col], default) if col else pd.Series(default, index=df.index, dtype="string")


def _month_series(df: pd.DataFrame) -> pd.Series:
    col = canonical_column(df, "DATE")
    return month_start(parse_dates(df[col])) if col else pd.Series(pd.NaT, index=df.index, dtype="datetime64[ns]")


def _qty_series(df: pd.DataFrame, preferred: Sequence[str] = ()) -> pd.Series:
    col = first_column(df, preferred, contains=False) if preferred else None
    col = col or canonical_column(df, "QTY")
    return quantity_mt(df[col], col) if col else pd.Series(0.0, index=df.index)


def _bu_forecast_source(sources: Mapping[str, SourceTable]) -> tuple[SourceTable | None, str]:
    """Prefer the BU-head UDT; retain the ESP export only as a fallback."""
    udt = _table_by_exact_tokens(sources, ["UDT_BU_HEAD_FCST", "BU_HEAD_FCST"])
    if udt is not None:
        return udt, "UDT_BU_HEAD_FCST"
    esp = _table_by_exact_tokens(sources, ["ESP_DFUTOSKUFCST_BU", "DFUTOSKUFCST_BU"])
    return esp, "ESP_DFUTOSKUFCST_BU fallback" if esp is not None else "Not found"


def _mapping_crosswalk(source: SourceTable | None) -> pd.DataFrame:
    columns = ["DFU", "DMDGROUP", "LOCATION", "MAPPED_ITEM", "BU"]
    if source is None or source.data.empty:
        return pd.DataFrame(columns=columns)
    df = source.data
    out = pd.DataFrame({
        "DFU": _id_series(df, "DFU"), "DMDGROUP": _dmdgroup_key(_id_series(df, "DMDGROUP")),
        "LOCATION": _id_series(df, "PLANT"), "MAPPED_ITEM": _id_series(df, "ITEM"),
        "BU": _normalise_bu(_id_series(df, "BU")),
    })
    blank = out["BU"].eq("(Blank)")
    out.loc[blank, "BU"] = _bu_from_item(out.loc[blank, "MAPPED_ITEM"])
    return out[columns].drop_duplicates().reset_index(drop=True)


def _source_date_bounds(source: SourceTable | None) -> tuple[pd.Timestamp | None, pd.Timestamp | None]:
    if source is None or source.data.empty:
        return None, None
    col = canonical_column(source.data, "DATE")
    if not col:
        return None, None
    dates = parse_dates(source.data[col]).dropna()
    return (dates.min(), dates.max()) if len(dates) else (None, None)


def _plan_horizon(sources: Mapping[str, SourceTable]) -> tuple[pd.Timestamp | None, pd.Timestamp | None, str]:
    """Use downstream SKU forecast as the authoritative comparison horizon."""
    sku_candidates = [s for s in _find_sources(sources, ["ESP_DFUTOSKUFCST", "DFUTOSKUFCST"]) if "_BU" not in s.name]
    for source in sku_candidates:
        start, end = _source_date_bounds(source)
        if start is not None:
            return month_start(pd.Series([start])).iloc[0], month_start(pd.Series([end])).iloc[0], source.path.name
    # Fall back to the explicit UDT horizon only when no downstream dates exist.
    source, _ = _bu_forecast_source(sources)
    start, end = _source_date_bounds(source)
    if start is not None:
        return month_start(pd.Series([start])).iloc[0], month_start(pd.Series([end])).iloc[0], source.path.name if source else ""
    return None, None, "No dated forecast source"


def _inventory_plan_start(sources: Mapping[str, SourceTable]) -> pd.Timestamp | None:
    """Earliest AVAILDATE in the latest non-parameter ESP_INVENTORY export."""
    candidates = [s for s in sources.values()
                  if "ESP_INVENTORY" in s.name and "PARAM" not in s.name and not s.data.empty]
    if not candidates:
        return None
    source = max(candidates, key=lambda s: s.path.stat().st_mtime_ns if s.path.exists() else 0)
    date_col = first_column(source.data, ("AVAILDATE", "AVAIL_DATE"))
    if not date_col:
        return None
    dates = parse_dates(source.data[date_col]).dropna()
    return pd.Timestamp(dates.min()).normalize() if len(dates) else None


def _resource_plan_end(sources: Mapping[str, SourceTable]) -> pd.Timestamp | None:
    """Last explicit weekly header in the latest RESOURCE_UTILIZATION matrix."""
    source = _latest_matching_source(sources, ("RESOURCE_UTILIZATION",))
    if source is None:
        return None
    dates = []
    for col in source.data.columns:
        parsed = parse_dates(pd.Series([col])).iloc[0]
        if pd.notna(parsed): dates.append(pd.Timestamp(parsed).normalize())
    return max(dates) if dates else None


def _forecast_key_reconciliation(sources: Mapping[str, SourceTable], source: SourceTable) -> dict[str, float | int]:
    cached = getattr(source, "_forecast_key_reconciliation_cache", None)
    if cached is not None:
        return cached
    sku_candidates = [s for s in _find_sources(sources, ["ESP_DFUTOSKUFCST", "DFUTOSKUFCST"]) if "_BU" not in s.name]
    if not sku_candidates:
        return {"matched_bu_raw": 0.0, "matched_sku_mt": 0.0, "matched_keys": 0, "bu_keys": 0, "matched_key_pct": 0.0}
    sdf = sku_candidates[0].data
    bdf = source.data
    bqty = first_column(bdf, ("SBU_FCST",)) or canonical_column(bdf, "QTY")
    sqty = first_column(sdf, FORECAST_MEASURES["DFUTOSKU_FCST"]) or canonical_column(sdf, "QTY")
    if not bqty or not sqty:
        return {"matched_bu_raw": 0.0, "matched_sku_mt": 0.0, "matched_keys": 0, "bu_keys": 0, "matched_key_pct": 0.0}
    def keyed(df: pd.DataFrame, values: pd.Series, name: str) -> pd.DataFrame:
        out = pd.DataFrame({
            "DFU": _id_series(df, "DFU"), "DMDGROUP": _dmdgroup_key(_id_series(df, "DMDGROUP")),
            "LOCATION": _id_series(df, "PLANT"), "MONTH": _month_series(df), name: values,
        })
        return out.groupby(["DFU", "DMDGROUP", "LOCATION", "MONTH"], dropna=False, observed=True, as_index=False)[name].sum()
    bu = keyed(bdf, numeric(bdf[bqty]), "BU_RAW")
    sku = keyed(sdf, quantity_mt(sdf[sqty], sqty), "SKU_MT")
    start, end, _ = _plan_horizon(sources)
    if start is not None:
        bu = bu[bu["MONTH"].between(start, end)]; sku = sku[sku["MONTH"].between(start, end)]
    bu_positive = bu[bu["BU_RAW"].gt(0)]
    joined = bu_positive.merge(sku, on=["DFU", "DMDGROUP", "LOCATION", "MONTH"], how="inner")
    result = {
        "matched_bu_raw": float(joined["BU_RAW"].sum()), "matched_sku_mt": float(joined["SKU_MT"].clip(lower=0).sum()),
        "matched_keys": int(len(joined)), "bu_keys": int(len(bu_positive)),
        "matched_key_pct": float(len(joined) / len(bu_positive) * 100) if len(bu_positive) else 0.0,
    }
    setattr(source, "_forecast_key_reconciliation_cache", result)
    return result


def _forecast_magnitude_stats(sources: Mapping[str, SourceTable], source: SourceTable) -> dict[str, Any]:
    """Compare fixed-unit BU and SKU totals over the downstream horizon."""
    cached = getattr(source, "_forecast_magnitude_cache", None)
    if cached is not None:
        return cached
    bdf = source.data; bqty = first_column(bdf, ("SBU_FCST",)) or canonical_column(bdf, "QTY")
    bu = pd.DataFrame({"MONTH": _month_series(bdf), "MT": numeric(bdf[bqty]) if bqty else 0.0})
    sku_source = _latest_matching_source(sources, ("ESP_DFUTOSKUFCST",))
    if sku_source is not None:
        sdf = sku_source.data; sqty = first_column(sdf, FORECAST_MEASURES["DFUTOSKU_FCST"]) or canonical_column(sdf, "QTY")
        sku = pd.DataFrame({"MONTH": _month_series(sdf), "MT": quantity_mt(sdf[sqty], sqty) if sqty else 0.0})
    else:
        sku = pd.DataFrame(columns=["MONTH", "MT"])
    start, end, _ = _plan_horizon(sources)
    if start is not None:
        bu = bu[bu["MONTH"].between(start, end)]; sku = sku[sku["MONTH"].between(start, end)]
    bu_total, sku_total = float(bu["MT"].clip(lower=0).sum()), float(sku["MT"].clip(lower=0).sum())
    ratio = bu_total / sku_total if sku_total > 0 else np.nan
    check = ("OK — same order of magnitude" if np.isfinite(ratio) and 0.1 <= ratio <= 10
             else "REVIEW — totals differ by more than 10x" if np.isfinite(ratio) else "REVIEW — SKU total unavailable")
    result = {"BU_DIVISOR": 1.0, "SKU_DIVISOR": 1000.0, "BU_TOTAL_MT": bu_total,
              "SKU_TOTAL_MT": sku_total, "MAGNITUDE_RATIO": ratio, "MAGNITUDE_CHECK": check}
    setattr(source, "_forecast_magnitude_cache", result)
    return result


def _bu_head_unit_scale(sources: Mapping[str, SourceTable], source: SourceTable) -> tuple[float, str, float, float]:
    """SBU_FCST is contractually already MT; never infer a kg divisor."""
    stats = _forecast_magnitude_stats(sources, source)
    basis = (f"Fixed divisors: BU SBU_FCST / {stats['BU_DIVISOR']:.0f}; "
             f"SKU TOTFCST / {stats['SKU_DIVISOR']:.0f}; BU/SKU={stats['MAGNITUDE_RATIO']:.3f} — {stats['MAGNITUDE_CHECK']}")
    return 1.0, basis, float(stats["BU_TOTAL_MT"]), float(stats["SKU_TOTAL_MT"])


def _derive_bu_head_rows(df: pd.DataFrame, mapping: pd.DataFrame, quantity_scale: float = 1.0) -> pd.DataFrame:
    """Normalise BU-head rows and derive BU/item from DFUTOSKU_FE."""
    out = pd.DataFrame({
        "DFU": _id_series(df, "DFU"), "DMDGROUP": _dmdgroup_key(_id_series(df, "DMDGROUP")),
        "LOCATION": _id_series(df, "PLANT"), "WEEK": week_start(_month_series(df)),
        "MONTH": _month_series(df), "BU": _normalise_bu(_id_series(df, "BU")),
        "BU_FCST_MT": numeric(df[first_column(df, ("SBU_FCST",)) or canonical_column(df, "QTY")]) * quantity_scale,
    })
    out["WEEK"] = week_start(parse_dates(df[canonical_column(df, "DATE")])) if canonical_column(df, "DATE") else pd.NaT
    out["MAPPED_ITEM"] = "(Blank)"
    if not mapping.empty:
        candidates = ["DFU", "DMDGROUP", "LOCATION"]
        join_keys = [k for k in candidates if not mapping[k].eq("(Blank)").all() and not out[k].eq("(Blank)").all()]
        join_keys = join_keys or ["DFU"]
        def first_nonblank(values: pd.Series) -> str:
            usable = values[~values.astype(str).isin(["", "(Blank)", "nan", "None"])]
            return str(usable.iloc[0]) if len(usable) else "(Blank)"
        mapped = (mapping.groupby(join_keys, dropna=False, observed=True, as_index=False)
                  .agg(MAPPED_ITEM=("MAPPED_ITEM", first_nonblank), BU=("BU", first_nonblank))
                  .rename(columns={"BU": "MAPPING_BU"}))
        out = out.merge(mapped, on=join_keys, how="left", suffixes=("", "_MAP"))
        if "MAPPED_ITEM_MAP" in out:
            out["MAPPED_ITEM"] = out["MAPPED_ITEM_MAP"].fillna(out["MAPPED_ITEM"]); out = out.drop(columns="MAPPED_ITEM_MAP")
        blank = out["BU"].eq("(Blank)")
        out.loc[blank, "BU"] = out.loc[blank, "MAPPING_BU"].fillna("(Blank)")
        out = out.drop(columns="MAPPING_BU")
    blank = out["BU"].eq("(Blank)")
    out.loc[blank, "BU"] = _bu_from_item(out.loc[blank, "MAPPED_ITEM"])
    # Some DFU names themselves carry the product prefix even if the mapping
    # is missing.  This is the last available deterministic fallback.
    blank = out["BU"].eq("(Blank)")
    out.loc[blank, "BU"] = _bu_from_item(out.loc[blank, "DFU"])
    return out


def build_dfu_mapping_loss(sources: Mapping[str, SourceTable]) -> pd.DataFrame:
    """Reconcile BU demand through FE mapping to supply forecast.

    The downstream ESP_DFUTOSKUFCST join is intentionally *only*
    supply ITEM + DFULOC + month.  DMDGROUP/DMDUNIT are demand-side
    attributes and must not be used to locate supply, per the corrected rule.
    """
    columns = ["DFU", "DMDGROUP", "LOC", "BU", "MONTH", "MAPPED_ITEM", "RATIO_SUM",
               "BU_FCST_MT", "EXPECTED_SKU_MT", "DFUTOSKU_FCST_MT", "LOST_MT", "REASON",
               "BU_DIVISOR", "SKU_DIVISOR", "MAGNITUDE_RATIO", "MAGNITUDE_CHECK"]
    bu_source, _ = _bu_forecast_source(sources)
    map_source = _latest_matching_source(sources, ("DFUTOSKU_FE",))
    sku_source = _latest_matching_source(sources, ("ESP_DFUTOSKUFCST",))
    if bu_source is None or map_source is None or sku_source is None:
        return pd.DataFrame(columns=columns)

    # Demand-side BU forecast.  Keep the scale inference because current
    # UDT_BU_HEAD_FCST exports can already be MT (for example 3-decimal files).
    bdf = bu_source.data
    bqty = first_column(bdf, ("SBU_FCST",)) or canonical_column(bdf, "QTY")
    scale, _, _, _ = _bu_head_unit_scale(sources, bu_source)
    bu = pd.DataFrame({
        "DMDGROUP": _dmdgroup_key(_id_series(bdf, "DMDGROUP")),
        "DFU": _id_series(bdf, "DFU"), "LOC": _id_series(bdf, "PLANT"),
        "MONTH": _month_series(bdf), "BU_FCST_MT": numeric(bdf[bqty]) * scale if bqty else 0.0,
    })
    horizon_start, horizon_end, _ = _plan_horizon(sources)
    if horizon_start is not None:
        bu = bu[bu["MONTH"].between(horizon_start, horizon_end)]
    bu = bu.groupby(["DMDGROUP", "DFU", "LOC", "MONTH"], dropna=False, observed=True, as_index=False)["BU_FCST_MT"].sum()

    # FE crosswalk: DMDUNIT (+ normalised DMDGROUP) + DFULOC -> supply ITEM.
    mdf = map_source.data
    m_dfu = canonical_column(mdf, "DFU"); m_item = canonical_column(mdf, "ITEM")
    m_loc = canonical_column(mdf, "PLANT"); m_group = canonical_column(mdf, "DMDGROUP")
    ratio_col = first_column(mdf, ("ALLOCFACTOR", "ALLOC_FACTOR", "RATIO", "SPLIT", "SPLITPCT",
                                   "ALLOCPCT", "PCT", "PERCENT", "DFUTOSKUFACTOR", "CONVFACTOR", "CONV_FACTOR"))
    if not m_dfu or not m_item or not m_loc:
        return pd.DataFrame(columns=columns)
    mapping = pd.DataFrame({
        "DFU": clean_text(mdf[m_dfu]), "LOC": clean_text(mdf[m_loc]),
        "DMDGROUP": _dmdgroup_key(clean_text(mdf[m_group])) if m_group else "*",
        "MAPPED_ITEM": clean_text(mdf[m_item]),
        "RATIO": numeric(mdf[ratio_col]) if ratio_col else 1.0,
    })
    mapping["RATIO"] = mapping["RATIO"].fillna(1.0)
    if ratio_col and ("PCT" in ratio_col or "PERCENT" in ratio_col or mapping["RATIO"].max() > 1.5):
        # Allocation-factor extracts sometimes store 50 for 50%; conversion
        # factors are only used as a fallback and are therefore not selected
        # when ALLOCFACTOR is present.
        mapping.loc[mapping["RATIO"].gt(1.5), "RATIO"] /= 100.0
    mapping = mapping.drop_duplicates()
    join_keys = ["DFU", "LOC"] + (["DMDGROUP"] if m_group else [])
    mapped = bu.merge(mapping, on=join_keys, how="left")
    mapped["HAS_MAPPING"] = mapped["MAPPED_ITEM"].notna()
    mapped["RATIO"] = mapped["RATIO"].fillna(0.0)
    mapped["EXPECTED_SKU_MT"] = mapped["BU_FCST_MT"] * mapped["RATIO"]

    # Supply-side forecast: exact corrected key = ITEM + DFULOC + month only.
    sdf = sku_source.data
    s_item = canonical_column(sdf, "ITEM"); s_loc = canonical_column(sdf, "PLANT")
    s_qty = first_column(sdf, FORECAST_MEASURES["DFUTOSKU_FCST"]) or canonical_column(sdf, "QTY")
    sku = pd.DataFrame({
        "MAPPED_ITEM": clean_text(sdf[s_item]) if s_item else "(Blank)",
        "LOC": clean_text(sdf[s_loc]) if s_loc else "(Blank)",
        "MONTH": _month_series(sdf),
        "SKU_MT": quantity_mt(sdf[s_qty], s_qty) if s_qty else 0.0,
        "BU": _normalise_bu(_id_series(sdf, "BU")),
    })
    sku = sku.groupby(["MAPPED_ITEM", "LOC", "MONTH"], dropna=False, observed=True, as_index=False).agg(
        SKU_MT=("SKU_MT", "sum"), BU=("BU", "first"))
    expected = (mapped[mapped["HAS_MAPPING"]]
                .groupby(["MAPPED_ITEM", "LOC", "MONTH"], dropna=False, observed=True)["EXPECTED_SKU_MT"]
                .sum().rename("EXPECTED_ITEM_MT").reset_index())
    mapped = mapped.merge(expected, on=["MAPPED_ITEM", "LOC", "MONTH"], how="left")
    mapped = mapped.merge(sku, on=["MAPPED_ITEM", "LOC", "MONTH"], how="left")
    mapped["SKU_MT"] = mapped["SKU_MT"].fillna(0.0)
    mapped["DFUTOSKU_FCST_MT"] = np.where(
        mapped["EXPECTED_ITEM_MT"].fillna(0).gt(0),
        mapped["SKU_MT"] * mapped["EXPECTED_SKU_MT"] / mapped["EXPECTED_ITEM_MT"], 0.0)
    sku_any = set(zip(sku["MAPPED_ITEM"], sku["LOC"]))
    mapped["SKU_ANY"] = [(i, l) in sku_any for i, l in zip(mapped["MAPPED_ITEM"], mapped["LOC"])]

    result = mapped.groupby(["DFU", "DMDGROUP", "LOC", "MONTH"], dropna=False, observed=True, as_index=False).agg(
        BU_FCST_MT=("BU_FCST_MT", "first"), HAS_MAPPING=("HAS_MAPPING", "any"),
        MAPPED_ITEM=("MAPPED_ITEM", lambda v: ", ".join(sorted(set(map(str, v.dropna())))[:3])),
        RATIO_SUM=("RATIO", "sum"), EXPECTED_SKU_MT=("EXPECTED_SKU_MT", "sum"),
        DFUTOSKU_FCST_MT=("DFUTOSKU_FCST_MT", "sum"), SKU_ANY=("SKU_ANY", "any"), BU=("BU", "first"))
    blank_bu = result["BU"].isna() | result["BU"].astype(str).isin(["", "(Blank)", "nan"])
    result.loc[blank_bu, "BU"] = _bu_from_item(result.loc[blank_bu, "MAPPED_ITEM"])
    result["LOST_MT"] = (result["BU_FCST_MT"] - result["DFUTOSKU_FCST_MT"]).clip(lower=0)
    result["REASON"] = np.select(
        [~result["HAS_MAPPING"], result["LOST_MT"].le(0.0005), ~result["SKU_ANY"], result["DFUTOSKU_FCST_MT"].le(0.0005)],
        ["No DFUTOSKU_FE mapping", "OK", "Mapped, item never in DFUTOSKUFCST", "Mapped, no SKU forecast this month"],
        default="Partial (SKU fcst < BU fcst)")
    magnitude = _forecast_magnitude_stats(sources, bu_source)
    for name in ("BU_DIVISOR", "SKU_DIVISOR", "MAGNITUDE_RATIO", "MAGNITUDE_CHECK"):
        result[name] = magnitude[name]
    result = result[result["LOST_MT"].gt(0.0005)].copy()
    return result[columns].sort_values("LOST_MT", ascending=False).reset_index(drop=True)


def build_bu_head_forecast(sources: Mapping[str, SourceTable]) -> pd.DataFrame:
    """Weekly DFU-level BU-head forecast used by the forecast funnel."""
    source, source_mode = _bu_forecast_source(sources)
    columns = ["WEEK", "MONTH", "DFU", "DMDGROUP", "LOCATION", "BU", "MAPPED_ITEM", "BU_FCST_MT", "DUR", "SOURCE_MODE", "SOURCE_FILE"]
    if source is None:
        return pd.DataFrame(columns=columns)
    mapping_source = _table_by_exact_tokens(sources, ["ESP_DFUTOSKU_FE", "DFUTOSKU_FE"])
    if "BU_HEAD_FCST" in source.name:
        scale, _, _, _ = _bu_head_unit_scale(sources, source)
        detail = _derive_bu_head_rows(source.data, _mapping_crosswalk(mapping_source), scale)
        dur_col = first_column(source.data, ("DUR", "DURATION", "BUCKETDURATION"))
        detail["DUR"] = clean_text(source.data[dur_col]) if dur_col else "(Blank)"
    else:
        df = source.data
        detail = pd.DataFrame({
            "WEEK": week_start(_month_series(df)), "MONTH": _month_series(df), "DFU": _id_series(df, "DFU"),
            "DMDGROUP": _id_series(df, "DMDGROUP"), "LOCATION": _id_series(df, "PLANT"),
            "BU": _normalise_bu(_id_series(df, "BU")), "MAPPED_ITEM": "(Blank)",
            "BU_FCST_MT": _qty_series(df, FORECAST_MEASURES["BU_FCST"]), "DUR": "(Blank)",
        })
    detail["SOURCE_MODE"] = source_mode; detail["SOURCE_FILE"] = source.path.name
    keys = ["WEEK", "MONTH", "DFU", "DMDGROUP", "LOCATION", "BU", "MAPPED_ITEM", "DUR", "SOURCE_MODE", "SOURCE_FILE"]
    return detail.groupby(keys, dropna=False, observed=True, as_index=False)["BU_FCST_MT"].sum()[columns].sort_values(["WEEK", "LOCATION", "DFU"]).reset_index(drop=True)


def _extract_forecast_measure(source: SourceTable, label: str) -> pd.DataFrame:
    df = source.data
    if df.empty: return pd.DataFrame()
    # ESP_INDDMDVIEW often holds FCST and CUSTORDER as row types.
    type_col = first_column(df, ("TYPE", "DMDTYPE", "DEMAND_TYPE", "ORDER_TYPE", "MEASURE"))
    if not type_col and "OPTIONSET" in df.columns:
        # OPTIONSET is a measure, never a simulation.  It is useful as a row
        # type only when its actual values identify forecast/order measures;
        # technical measure names such as SNOPWI_* must not filter all rows out.
        option_values = clean_text(df["OPTIONSET"], "").str.upper()
        if option_values.str.contains("FCST|FORECAST|CUST|ORDER", regex=True).any():
            type_col = "OPTIONSET"
    qty_col = first_column(df, FORECAST_MEASURES[label]) or canonical_column(df, "QTY")
    if not qty_col: return pd.DataFrame()
    mask = pd.Series(True, index=df.index)
    if type_col and label in {"NETTED_FCST", "CUSTORDER"}:
        values = clean_text(df[type_col], "").str.upper()
        if label == "NETTED_FCST": mask = values.str.contains("FCST|FORECAST", regex=True)
        else: mask = values.str.contains("CUST|ORDER", regex=True)
    dims = dimension_frame(df.loc[mask])
    dims["DMDGROUP"] = _dmdgroup_key(dims["DMDGROUP"])
    dims["MONTH"] = _month_series(df.loc[mask])
    dims[label + "_MT"] = quantity_mt(df.loc[mask, qty_col], qty_col)
    return dims


def build_forecast_flow(sources: Mapping[str, SourceTable]) -> pd.DataFrame:
    keys = ["MONTH", "BU", "PLANT", "ITEM", "DFU", "DMDGROUP"]
    pieces: dict[str, list[pd.DataFrame]] = {k: [] for k in FORECAST_MEASURES}
    preferred_bu_source, _ = _bu_forecast_source(sources)
    # Enrich BU-level forecast with the mapped SKU where the mapping exists so
    # the end-to-end stages line up on one row instead of creating a blank-ITEM
    # BU row and a separate SKU row.
    mapping_lookup = pd.DataFrame()
    mapping_source = _table_by_exact_tokens(sources, ["ESP_DFUTOSKU_FE", "DFUTOSKU_FE"])
    if mapping_source is not None and not mapping_source.data.empty:
        md = _mapping_crosswalk(mapping_source).rename(columns={"LOCATION": "PLANT", "MAPPED_ITEM": "ITEM"})
        mapping_lookup = md[["BU", "PLANT", "ITEM", "DFU", "DMDGROUP"]].drop_duplicates()
    for source in sources.values():
        n = source.name
        labels: list[str] = []
        if "UDT_BU_HEAD_FCST" in n or "BU_HEAD_FCST" in n:
            labels = ["BU_FCST"] if source is preferred_bu_source else []
        elif "DFUTOSKUFCST_BU" in n:
            labels = ["BU_FCST"] if source is preferred_bu_source else []
        elif "DFUTOSKUFCST" in n: labels = ["DFUTOSKU_FCST"]
        elif "INDDMDVIEW" in n: labels = ["NETTED_FCST", "CUSTORDER"]
        elif "CUSTORDERSTATUS" in n: labels = []
        else:
            for label, aliases in FORECAST_MEASURES.items():
                if any(a in n for a in aliases): labels.append(label)
        for label in labels:
            if label == "BU_FCST" and ("UDT_BU_HEAD_FCST" in n or "BU_HEAD_FCST" in n):
                scale, _, _, _ = _bu_head_unit_scale(sources, source)
                detail = _derive_bu_head_rows(source.data, _mapping_crosswalk(mapping_source), scale)
                horizon_start, horizon_end, _ = _plan_horizon(sources)
                if horizon_start is not None:
                    detail = detail[detail["MONTH"].between(horizon_start, horizon_end)]
                piece = pd.DataFrame({
                    "MONTH": detail["MONTH"], "BU": detail["BU"], "PLANT": detail["LOCATION"],
                    "ITEM": detail["MAPPED_ITEM"], "DFU": detail["DFU"], "DMDGROUP": detail["DMDGROUP"],
                    "BU_FCST_MT": detail["BU_FCST_MT"],
                })
                piece["RESOURCE"] = "(Blank)"; piece["PRODUCT_FORM"] = "(Blank)"; piece["STREAM"] = "(Blank)"; piece["SUB_CLASS"] = "(Blank)"; piece["LOCATION_TYPE"] = np.where(piece["PLANT"].str.upper().str.startswith("V_"), "Demand zone", "Plant")
            else:
                piece = _extract_forecast_measure(source, label)
            if label == "BU_FCST" and not piece.empty and not mapping_lookup.empty:
                candidate_keys = ["DFU", "DMDGROUP", "BU", "PLANT"]
                join_keys = [k for k in candidate_keys if not mapping_lookup[k].eq("(Blank)").all()]
                join_keys = join_keys or ["DFU"]
                lookup = mapping_lookup[[*join_keys, "ITEM"]].drop_duplicates(join_keys).rename(columns={"ITEM": "MAPPED_ITEM"})
                piece = piece.merge(lookup, on=join_keys, how="left")
                piece["ITEM"] = piece["ITEM"].mask(piece["ITEM"].eq("(Blank)"), piece["MAPPED_ITEM"]).fillna("(Blank)")
                piece = piece.drop(columns="MAPPED_ITEM")
            elif label in {"DFUTOSKU_FCST", "SCHED_QTY"} and not piece.empty and not mapping_lookup.empty:
                join_keys = [k for k in ["ITEM", "PLANT", "BU"] if not mapping_lookup[k].eq("(Blank)").all() and not piece[k].eq("(Blank)").all()]
                if join_keys:
                    lookup = mapping_lookup[[*join_keys, "DFU", "DMDGROUP"]].drop_duplicates(join_keys).rename(columns={"DFU": "MAPPED_DFU", "DMDGROUP": "MAPPED_DMDGROUP"})
                    piece = piece.merge(lookup, on=join_keys, how="left")
                    piece["DFU"] = piece["DFU"].mask(piece["DFU"].eq("(Blank)"), piece["MAPPED_DFU"]).fillna("(Blank)")
                    piece["DMDGROUP"] = piece["DMDGROUP"].mask(piece["DMDGROUP"].eq("(Blank)"), piece["MAPPED_DMDGROUP"]).fillna("(Blank)")
                    piece = piece.drop(columns=["MAPPED_DFU", "MAPPED_DMDGROUP"])
            if not piece.empty: pieces[label].append(piece)
    result: pd.DataFrame | None = None
    for label in FORECAST_MEASURES:
        if not pieces[label]: continue
        part = pd.concat(pieces[label], ignore_index=True)
        value = label + "_MT"
        part = part.groupby(keys, dropna=False, observed=True, as_index=False)[value].sum()
        result = part if result is None else result.merge(part, on=keys, how="outer")
    if result is None:
        return pd.DataFrame(columns=[*keys, *[x + "_MT" for x in FORECAST_MEASURES], "MET_PCT"])
    for label in FORECAST_MEASURES:
        col = label + "_MT"
        if col not in result: result[col] = 0.0
        result[col] = result[col].fillna(0.0)
    # Derive TOTAL_FCST if no explicit export exists.
    if result["TOTAL_FCST_MT"].abs().sum() == 0:
        result["TOTAL_FCST_MT"] = result["NETTED_FCST_MT"] + result["CUSTORDER_MT"]
    result["MET_PCT"] = np.where(result["TOTAL_FCST_MT"].gt(0), result["SCHED_QTY_MT"] / result["TOTAL_FCST_MT"] * 100, np.nan)
    return result.sort_values(["MONTH", "BU", "PLANT", "ITEM"], kind="stable").reset_index(drop=True)


def build_order_priority(sources: Mapping[str, SourceTable]) -> pd.DataFrame:
    frames = [_fact_from_source(s) for s in sources.values() if s.kind == "order" or "CUSTORDER" in s.name]
    if not frames: return pd.DataFrame(columns=["DATE", "PLANT", "ITEM", "CUSTOMER", "ORDER_ID", "PRIORITY", "QTY_MT"])
    out = pd.concat(frames, ignore_index=True)
    cols = ["DATE", "BU", "PLANT", "ITEM", "CUSTOMER", "ORDER_ID", "PRIORITY", "QTY_MT", "SOURCE_FILE"]
    return out[cols].sort_values(["PRIORITY", "DATE", "QTY_MT"], ascending=[True, True, False], kind="stable").reset_index(drop=True)


def build_plant_cards(trend: pd.DataFrame, flow: pd.DataFrame) -> pd.DataFrame:
    if trend.empty and flow.empty: return pd.DataFrame()
    if trend.empty:
        cards = flow.groupby(["BU", "PLANT"], dropna=False, observed=True, as_index=False)[["TOTAL_FCST_MT", "SCHED_QTY_MT"]].sum()
    else:
        flow_measures = trend[~trend["MEASURE"].isin(["Opening Inventory", "Closing Inventory"])]
        pivot = (flow_measures.groupby(["BU", "PLANT", "MEASURE"], dropna=False, observed=True)["QTY_MT"].sum()
                 .unstack("MEASURE", fill_value=0).reset_index()) if not flow_measures.empty else trend[["BU", "PLANT"]].drop_duplicates()
        pivot.columns.name = None
        opening = (trend[trend["MEASURE"].eq("Opening Inventory")].sort_values("WEEK")
                   .groupby(["BU", "PLANT"], dropna=False, observed=True, as_index=False).first()[["BU", "PLANT", "QTY_MT"]]
                   .rename(columns={"QTY_MT": "Opening Inventory"}))
        closing = (trend[trend["MEASURE"].eq("Closing Inventory")].sort_values("WEEK")
                   .groupby(["BU", "PLANT"], dropna=False, observed=True, as_index=False).last()[["BU", "PLANT", "QTY_MT"]]
                   .rename(columns={"QTY_MT": "Closing Inventory"}))
        if not opening.empty: pivot = pivot.merge(opening, on=["BU", "PLANT"], how="outer")
        if not closing.empty: pivot = pivot.merge(closing, on=["BU", "PLANT"], how="outer")
        cards = pivot
        if not flow.empty:
            f = flow.groupby(["BU", "PLANT"], dropna=False, observed=True, as_index=False)[["TOTAL_FCST_MT", "SCHED_QTY_MT"]].sum()
            cards = cards.merge(f, on=["BU", "PLANT"], how="outer")
    for col in [*MEASURE_ALIASES.keys(), "TOTAL_FCST_MT", "SCHED_QTY_MT"]:
        if col not in cards: cards[col] = 0.0
        cards[col] = cards[col].fillna(0.0)
    cards["LOCATION_TYPE"] = np.where(clean_text(cards["PLANT"], "").str.upper().str.startswith("V_"), "Demand zone", "Plant")
    cards["ROUTED_SUPPLY_MT"] = np.where(cards["LOCATION_TYPE"].eq("Demand zone"), cards["In-Transit"], cards["SCHED_QTY_MT"])
    effective_supply = np.where(cards["LOCATION_TYPE"].eq("Demand zone"), np.maximum(cards["SCHED_QTY_MT"], cards["ROUTED_SUPPLY_MT"]), cards["SCHED_QTY_MT"])
    cards["SUPPLY_GAP_MT"] = (cards["TOTAL_FCST_MT"] - effective_supply).clip(lower=0)
    cards["SERVICE_PCT"] = np.where(cards["TOTAL_FCST_MT"].gt(0), effective_supply / cards["TOTAL_FCST_MT"] * 100, np.nan)
    return cards.sort_values("SUPPLY_GAP_MT", ascending=False).reset_index(drop=True)


def build_schema_report(sources: Mapping[str, SourceTable]) -> pd.DataFrame:
    rows = []
    for source in sources.values():
        matched = {field: canonical_column(source.data, field) or "" for field in ("DATE", "BU", "PLANT", "RESOURCE", "ITEM", "DFU", "DMDGROUP", "QTY")}
        is_bu_head = "UDT_BU_HEAD_FCST" in source.name or "BU_HEAD_FCST" in source.name
        rows.append({"FILE": source.path.name, "FILE_GROUP": "Input UDT" if is_bu_head else "ESP / UDT source",
                     "CATALOGUE_ROLE": "BU head forecast" if is_bu_head else source.kind.replace("_", " ").title(),
                     "CLASS": source.kind, "ROWS_BEFORE_SIMULATION": source.rows_before_filter, "ROWS": len(source.data),
                     "ROWS_REMOVED": source.rows_before_filter-len(source.data), "LOAD_ERROR": source.load_error,
                     "COLUMNS": len(source.data.columns), "SIMULATION_FILTER": source.simulation_note,
                     "AVAILABLE_SIMULATIONS": ", ".join(source.available_simulations), **{f"MATCHED_{k}": v for k, v in matched.items()},
                     "COLUMN_LIST": ", ".join(source.data.columns)})
    return pd.DataFrame(rows)


def build_mapping_matches(sources: Mapping[str, SourceTable]) -> pd.DataFrame:
    rows = []
    preferred_bu, bu_source_type = _bu_forecast_source(sources)
    specs = [
        ("BU forecast", ["UDT_BU_HEAD_FCST", "BU_HEAD_FCST"] if preferred_bu and "BU_HEAD_FCST" in preferred_bu.name else ["ESP_DFUTOSKUFCST_BU", "DFUTOSKUFCST_BU"], FORECAST_MEASURES["BU_FCST"]),
        ("SKU forecast", ["ESP_DFUTOSKUFCST", "DFUTOSKUFCST"], FORECAST_MEASURES["DFUTOSKU_FCST"]),
        ("DFU-to-SKU mapping", ["ESP_DFUTOSKU_FE", "DFUTOSKU_FE"], ("CONVFACTOR", "FACTOR", "QTY")),
    ]
    for role, tokens, qty_aliases in specs:
        candidates = _find_sources(sources, tokens)
        if role == "SKU forecast": candidates = [s for s in candidates if "_BU" not in s.name]
        source = candidates[0] if candidates else None
        df = source.data if source else pd.DataFrame()
        rows.append({
            "ROLE": role, "SOURCE_MODE": bu_source_type if role == "BU forecast" else "Required downstream stage",
            "FILE": source.path.name if source else "Not found", "ROWS": len(df),
            "DFU_COLUMN": (canonical_column(df, "DFU") or "") if not df.empty else "",
            "DMDGROUP_COLUMN": (canonical_column(df, "DMDGROUP") or "") if not df.empty else "",
            "LOCATION_COLUMN": (canonical_column(df, "PLANT") or "") if not df.empty else "",
            "BU_COLUMN": (canonical_column(df, "BU") or "") if not df.empty else "",
            "ITEM_COLUMN": (canonical_column(df, "ITEM") or "") if not df.empty else "",
            "MONTH_COLUMN": (canonical_column(df, "DATE") or "") if not df.empty else "",
            "QUANTITY_COLUMN": ((first_column(df, qty_aliases) or canonical_column(df, "QTY") or "") if not df.empty else ""),
        })
    return pd.DataFrame(rows)


def build_forecast_reconciliation(sources: Mapping[str, SourceTable], flow: pd.DataFrame, loss: pd.DataFrame) -> pd.DataFrame:
    source, source_mode = _bu_forecast_source(sources)
    start, end, horizon_source = _plan_horizon(sources)
    scale, unit_basis, raw_total, reference_mt = ((1.0, "No BU forecast", 0.0, 0.0)
                                                   if source is None else _bu_head_unit_scale(sources, source))
    magnitude = ({"BU_DIVISOR": 1.0, "SKU_DIVISOR": 1000.0, "BU_TOTAL_MT": 0.0, "SKU_TOTAL_MT": 0.0,
                  "MAGNITUDE_RATIO": np.nan, "MAGNITUDE_CHECK": "No BU forecast"}
                 if source is None else _forecast_magnitude_stats(sources, source))
    loss_by_reason = loss.groupby("REASON", observed=True)["LOST_MT"].sum().to_dict() if not loss.empty else {}
    metrics = {
        "BU source": source.path.name if source else "Not found",
        "BU source mode": source_mode,
        "SBU_FCST raw total": raw_total,
        "SBU_FCST scale to MT": scale,
        "Unit decision": unit_basis,
        "Reference total MT": reference_mt,
        "BU divisor (SBU_FCST already MT)": magnitude["BU_DIVISOR"],
        "SKU divisor (TOTFCST kg to MT)": magnitude["SKU_DIVISOR"],
        "BU total MT in horizon": magnitude["BU_TOTAL_MT"],
        "SKU total MT in horizon": magnitude["SKU_TOTAL_MT"],
        "BU/SKU magnitude ratio": magnitude["MAGNITUDE_RATIO"],
        "Magnitude check": magnitude["MAGNITUDE_CHECK"],
        "Plan horizon start": start,
        "Plan horizon end": end,
        "Plan horizon source": horizon_source,
        "BU_FCST_MT in funnel": float(flow.get("BU_FCST_MT", pd.Series(dtype=float)).sum()),
        "DFUTOSKU_FCST_MT in funnel": float(flow.get("DFUTOSKU_FCST_MT", pd.Series(dtype=float)).sum()),
        "Lost MT": float(loss.get("LOST_MT", pd.Series(dtype=float)).sum()),
        "Loss rows": int(len(loss)),
        "Lost MT — no mapping": float(loss_by_reason.get("No DFUTOSKU mapping", 0.0)),
        "Lost MT — mapped zero": float(loss_by_reason.get("Mapped but zero SKU forecast", 0.0)),
        "Lost MT — partial": float(loss_by_reason.get("Partial", 0.0)),
    }
    return pd.DataFrame({"CHECK": list(metrics), "VALUE": list(metrics.values())})


def build_service_demand(sources: Mapping[str, SourceTable]) -> pd.DataFrame:
    frames = []
    preferred = _latest_matching_source(sources, ("UDT_ORDER_FULFILLMENT", "ORDER_FULFILLMENT"))
    for source in sources.values():
        if "UDT_ORDER_FULFILLMENT" not in source.name and "ORDER_FULFILLMENT" not in source.name:
            continue
        if preferred is not None and source is not preferred: continue
        df = source.data
        if df.empty: continue
        dims = dimension_frame(df)
        date_col = first_column(df, ("MONTH", "NEEDDATE", "DUEDATE", "ORDERDATE", "DATE"))
        demand_col = first_column(df, ("TOTAL_ORDER", "TOTALORDER", "TOTAL_CUST_ORDER", "TOTALCUSTORDER", "DEMANDQTY", "ORDERQTY", "DMDQTY", "TOTALDEMAND", "QTY"))
        allocated_col = first_column(df, ("TOTAL_ALLOCATION", "TOTALALLOCATION", "TOT_CUSTORDER_ALLOCATION", "TOTCUSTORDERALLOCATION", "ALLOCATEDQTY", "FULFILLEDQTY", "SCHEDULEDQTY", "SCHEDQTY", "SUPPLYQTY", "METQTY"))
        unallocated_col = first_column(df, ("UNALLOCATEDQTY", "UNALLOCQTY", "SHORTQTY", "GAPQTY", "UNMETQTY"))
        if not demand_col: continue
        out = dims.copy(); out["DATE"] = parse_dates(df[date_col]) if date_col else pd.NaT
        out["DEMAND_MT"] = quantity_mt(df[demand_col], demand_col)
        out["ALLOCATED_MT"] = quantity_mt(df[allocated_col], allocated_col) if allocated_col else 0.0
        out["UNALLOCATED_MT"] = quantity_mt(df[unallocated_col], unallocated_col) if unallocated_col else (out["DEMAND_MT"] - out["ALLOCATED_MT"]).clip(lower=0)
        out["ORDER_FILL_PCT"] = np.where(out["DEMAND_MT"].gt(0), out["ALLOCATED_MT"] / out["DEMAND_MT"] * 100, np.nan)
        cust_demand = first_column(df, ("TOTAL_CUST_ORDER", "TOTALCUSTORDER"))
        cust_alloc = first_column(df, ("TOT_CUSTORDER_ALLOCATION", "TOTCUSTORDERALLOCATION"))
        fcst_demand = first_column(df, ("TOT_FORCAST_ORDER", "TOT_FORECAST_ORDER", "TOTAL_FORECAST_ORDER"))
        fcst_alloc = first_column(df, ("TOT_FCSTORDER_ALLOCATION", "TOT_FORECASTORDER_ALLOCATION"))
        out["CUST_ORDER_MT"] = quantity_mt(df[cust_demand], cust_demand) if cust_demand else 0.0
        out["CUST_ALLOCATION_MT"] = quantity_mt(df[cust_alloc], cust_alloc) if cust_alloc else 0.0
        out["FCST_ORDER_MT"] = quantity_mt(df[fcst_demand], fcst_demand) if fcst_demand else 0.0
        out["FCST_ALLOCATION_MT"] = quantity_mt(df[fcst_alloc], fcst_alloc) if fcst_alloc else 0.0
        out["MATCHED_DEMAND_COLUMN"] = demand_col; out["MATCHED_ALLOCATION_COLUMN"] = allocated_col or ""
        out["MATCHED_DATE_COLUMN"] = date_col or ""
        out["SOURCE_FILE"] = source.path.name
        out = out.loc[~out["BU"].astype(str).isin(["", "(Blank)", "(BLANK)", "NAN", "NONE"])]
        frames.append(out)
    columns = ["DATE", "BU", "PLANT", "LOCATION_TYPE", "ITEM", "DFU", "DMDGROUP", "DEMAND_MT", "ALLOCATED_MT", "UNALLOCATED_MT", "ORDER_FILL_PCT",
               "CUST_ORDER_MT", "CUST_ALLOCATION_MT", "FCST_ORDER_MT", "FCST_ALLOCATION_MT", "MATCHED_DEMAND_COLUMN", "MATCHED_ALLOCATION_COLUMN", "MATCHED_DATE_COLUMN", "SOURCE_FILE"]
    if not frames: return pd.DataFrame(columns=columns)
    return pd.concat(frames, ignore_index=True)[columns].sort_values("UNALLOCATED_MT", ascending=False).reset_index(drop=True)


def build_value_at_risk(sources: Mapping[str, SourceTable]) -> pd.DataFrame:
    frames = []
    for source in sources.values():
        if "ESP_INDDMDVIEW" not in source.name and "INDDMDVIEW" not in source.name:
            continue
        df = source.data
        margin_col = first_column(df, ("CALCMARGIN", "CALC_MARGIN", "MARGIN", "CONTRIBUTIONMARGIN"))
        if not margin_col: continue
        dims = dimension_frame(df); date_col = canonical_column(df, "DATE")
        sched_col = first_column(df, ("SCHEDQTY", "SCHED_QTY", "SCHEDULEDQTY", "ALLOCATEDQTY", "SUPPLYQTY", "QTY"))
        demand_col = first_column(df, ("DMDQTY", "DEMANDQTY", "ORDERQTY", "FCSTQTY", "QTY"))
        out = dims.copy(); out["DATE"] = parse_dates(df[date_col]) if date_col else pd.NaT
        out["CALCMARGIN"] = numeric(df[margin_col])
        out["DEMAND_MT"] = quantity_mt(df[demand_col], demand_col) if demand_col else 0.0
        out["SCHEDULED_MT"] = quantity_mt(df[sched_col], sched_col) if sched_col else 0.0
        out["NEGATIVE_MARGIN_SCHEDULED_MT"] = np.where(out["CALCMARGIN"].lt(0), out["SCHEDULED_MT"], 0.0)
        out["VALUE_AT_RISK"] = np.where(out["CALCMARGIN"].lt(0), -out["CALCMARGIN"], 0.0)
        out["SOURCE_FILE"] = source.path.name
        frames.append(out)
    columns = ["DATE", "BU", "PLANT", "LOCATION_TYPE", "ITEM", "DFU", "DMDGROUP", "CALCMARGIN", "DEMAND_MT", "SCHEDULED_MT", "NEGATIVE_MARGIN_SCHEDULED_MT", "VALUE_AT_RISK", "SOURCE_FILE"]
    if not frames: return pd.DataFrame(columns=columns)
    return pd.concat(frames, ignore_index=True)[columns].sort_values(["VALUE_AT_RISK", "NEGATIVE_MARGIN_SCHEDULED_MT"], ascending=False).reset_index(drop=True)


def _measure_value_rows(df: pd.DataFrame, labels: Sequence[str]) -> tuple[pd.Series, str | None, str | None]:
    label_set = {normalise_name(x) for x in labels}
    for col in df.columns:
        vals = clean_text(df[col], "").map(normalise_name)
        mask = vals.isin(label_set)
        if mask.any():
            value_col = _find_value_column(df.loc[mask], ("VALUE", "QTY", "UTILIZATION", "PERCENT", "PCT"), excluded=[col])
            return mask, col, value_col
    return pd.Series(False, index=df.index), None, None


def build_capacity(sources: Mapping[str, SourceTable]) -> pd.DataFrame:
    frames = []
    preferred_resource_util = _latest_matching_source(sources, ("RESOURCE_UTILIZATION",))
    shutdown_keys: set[tuple[str, str]] = set()
    for source in sources.values():
        if "SHUTDOWN" in source.name or "RESCAL" in source.name or "RESOURCE_CALENDAR" in source.name:
            d = source.data; res = canonical_column(d, "RESOURCE"); loc = canonical_column(d, "PLANT")
            if res:
                shutdown_keys.update(zip(clean_text(d[loc]) if loc else pd.Series("(Blank)", index=d.index), clean_text(d[res])))
    for source in sources.values():
        if not any(token in source.name for token in ("RESOURCE_UTILIZATION", "UDT_RES_UTIL_AND_ALLOC", "ESP_RESLOADDETAIL", "RESOURCE_CAPACITY")):
            continue
        df = source.data
        if df.empty: continue
        # RESOURCE_UTILIZATION is a wide weekly matrix. Normalised headers such
        # as 22_OCT_26_12_00_AM are parsed back to their bucket dates.
        if "RESOURCE_UTILIZATION" in source.name:
            if preferred_resource_util is not None and source is not preferred_resource_util: continue
            date_columns: list[str] = []
            parsed_headers: dict[str, pd.Timestamp] = {}
            for col in df.columns:
                parsed = parse_dates(pd.Series([col])).iloc[0]
                if pd.notna(parsed): date_columns.append(col); parsed_headers[col] = parsed
            resource_col = canonical_column(df, "RESOURCE")
            plant_col = canonical_column(df, "PLANT")
            bu_col = canonical_column(df, "BU")
            option_col = first_column(df, ("OPTIONSET", "OPTION_SET", "MEASURE"))
            if date_columns and resource_col:
                base = df.copy()
                if option_col:
                    option_values = clean_text(base[option_col], "").str.upper()
                    util_rows = option_values.str.contains("RESOURCE.*UTIL|%.*UTIL", regex=True)
                    if util_rows.any(): base = base.loc[util_rows]
                resource = clean_text(base[resource_col])
                keep = (~resource.str.upper().str.contains("DUMMY", regex=False)
                        & ~resource.str.upper().isin(["", "(BLANK)", "NAN", "NONE", "TESTSBUY"]))
                base = base.loc[keep]
                id_cols = [c for c in base.columns if c not in date_columns]
                melted = base.melt(id_vars=id_cols, value_vars=date_columns, var_name="WEEK_HEADER", value_name="UTIL_VALUE")
                melted["DATE"] = melted["WEEK_HEADER"].map(parsed_headers)
                util = numeric(melted["UTIL_VALUE"])
                # RESOURCE_UTILIZATION's PCTUSED matrix is already expressed
                # in percentage points.  Do not rescale isolated 0/1 values;
                # doing so incorrectly added SIL_ANOD as a 45th bottleneck.
                melted["UTILIZATION_PCT"] = util
                melted["PLANT"] = clean_text(melted[plant_col]) if plant_col else "(Blank)"
                melted["RESOURCE"] = clean_text(melted[resource_col])
                melted["BU"] = _normalise_bu(clean_text(melted[bu_col])) if bu_col else "(Blank)"
                melted["STREAM"] = "(Blank)"; melted["CAPACITY_MT"] = np.nan; melted["LOAD_MT"] = np.nan
                melted["SHUTDOWN_CALENDAR"] = [(p, r) in shutdown_keys or ("(Blank)", r) in shutdown_keys for p, r in zip(melted["PLANT"], melted["RESOURCE"])]
                melted["BOTTLENECK"] = melted["UTILIZATION_PCT"].ge(95)
                melted["SOURCE_FILE"] = source.path.name
                frames.append(melted[["DATE", "BU", "PLANT", "RESOURCE", "STREAM", "CAPACITY_MT", "LOAD_MT", "UTILIZATION_PCT", "SHUTDOWN_CALENDAR", "BOTTLENECK", "SOURCE_FILE"]])
                continue
        dims = dimension_frame(df); date_col = canonical_column(df, "DATE")
        util_col = first_column(df, ("UTILIZATION_PCT", "UTILIZATION", "RESOURCE_UTIL", "RESUTIL", "PERCENT_UTILIZATION", "PCT_UTIL"))
        cap_col = canonical_column(df, "CAPACITY")
        load_col = canonical_column(df, "LOAD")
        row_mask = pd.Series(True, index=df.index)
        if not util_col and "UDT_RES_UTIL_AND_ALLOC" in source.name:
            row_mask, _, util_col = _measure_value_rows(df, ("% Resource Util", "Resource Util", "Resource Utilization"))
        out = dims.loc[row_mask].copy()
        out["DATE"] = parse_dates(df.loc[row_mask, date_col]) if date_col else pd.NaT
        out["CAPACITY_MT"] = quantity_mt(df.loc[row_mask, cap_col], cap_col) if cap_col else np.nan
        out["LOAD_MT"] = quantity_mt(df.loc[row_mask, load_col], load_col) if load_col else np.nan
        if util_col:
            util = numeric(df.loc[row_mask, util_col]); out["UTILIZATION_PCT"] = np.where(util.abs().le(1.5), util * 100, util)
        else:
            out["UTILIZATION_PCT"] = np.where(out["CAPACITY_MT"].gt(0), out["LOAD_MT"] / out["CAPACITY_MT"] * 100, np.nan)
        out["SHUTDOWN_CALENDAR"] = [(p, r) in shutdown_keys or ("(Blank)", r) in shutdown_keys for p, r in zip(out["PLANT"], out["RESOURCE"])]
        out["BOTTLENECK"] = out["UTILIZATION_PCT"].ge(95)
        out["SOURCE_FILE"] = source.path.name
        frames.append(out)
    columns = ["DATE", "BU", "PLANT", "RESOURCE", "STREAM", "CAPACITY_MT", "LOAD_MT", "UTILIZATION_PCT", "SHUTDOWN_CALENDAR", "BOTTLENECK", "SOURCE_FILE"]
    if not frames: return pd.DataFrame(columns=columns)
    out = pd.concat(frames, ignore_index=True)[columns]
    out["SHUTDOWN_CALENDAR"] = out["SHUTDOWN_CALENDAR"].fillna(False).astype(bool)
    out["BOTTLENECK"] = out["BOTTLENECK"].fillna(False).astype(bool)
    keys = ["DATE", "BU", "PLANT", "RESOURCE", "STREAM"]
    authoritative = out["SOURCE_FILE"].astype(str).str.contains("RESOURCE_UTILIZATION", case=False, na=False)
    if authoritative.any():
        base = out.loc[authoritative & out["UTILIZATION_PCT"].notna()].copy()
        if base.empty:
            return pd.DataFrame(columns=columns)
        # Keep resources that have a shutdown-calendar record in the denominator.
        # The calendar is an annotation, not a reason to remove a machine centre;
        # dropping these rows caused the v2.4.0 regression from 44/133 to 35/121.
        resource_keys = ["BU", "PLANT", "RESOURCE", "STREAM"]
        idx = base.groupby(resource_keys, dropna=False, observed=True)["UTILIZATION_PCT"].idxmax()
        out = base.loc[idx].copy()
        counts = (base.assign(IS_HIGH_98=base["UTILIZATION_PCT"].ge(98))
                  .groupby(resource_keys, dropna=False, observed=True, as_index=False)
                  .agg(HIGH_98_WEEKS=("IS_HIGH_98", "sum"), WEEKS_WITH_DATA=("UTILIZATION_PCT", "count"),
                       MIN_UTILIZATION_PCT=("UTILIZATION_PCT", "min"), AVG_UTILIZATION_PCT=("UTILIZATION_PCT", "mean")))
        out = out.merge(counts, on=resource_keys, how="left")
        high_share = np.where(out["WEEKS_WITH_DATA"].gt(0), out["HIGH_98_WEEKS"] / out["WEEKS_WITH_DATA"], 0.0)
        out["BOTTLENECK"] = (high_share >= .5) | out["AVG_UTILIZATION_PCT"].ge(95)
        out["RESOURCE_STATUS"] = np.select(
            [out["BOTTLENECK"], out["AVG_UTILIZATION_PCT"].lt(30), out["AVG_UTILIZATION_PCT"].ge(85) | (high_share >= .25)],
            ["Bottleneck", "Under-used", "Tight"], default="Normal")
    else:
        out = out.sort_values("UTILIZATION_PCT", ascending=False).drop_duplicates(keys, keep="first")
        out["BOTTLENECK"] = out["UTILIZATION_PCT"].ge(95)
        out["AVG_UTILIZATION_PCT"] = out["UTILIZATION_PCT"]
        out["HIGH_98_WEEKS"] = out["UTILIZATION_PCT"].ge(98).astype(int)
        out["WEEKS_WITH_DATA"] = out["UTILIZATION_PCT"].notna().astype(int)
        out["RESOURCE_STATUS"] = np.select([out["BOTTLENECK"], out["UTILIZATION_PCT"].lt(30), out["UTILIZATION_PCT"].ge(85)], ["Bottleneck", "Under-used", "Tight"], default="Normal")

    # Resource-level output is the available load proxy when the utilization
    # matrix has no explicit load/capacity quantities.
    load_parts = []
    for source in sources.values():
        if "UDT_RES_UTIL_AND_ALLOC" in source.name and not source.data.empty:
            part = _resource_output_fact(source)
            if not part.empty:
                load_parts.append(part.groupby(["PLANT", "RESOURCE"], dropna=False, observed=True, as_index=False)["QTY_MT"].sum().rename(columns={"QTY_MT": "OUTPUT_LOAD_MT"}))
    if load_parts:
        load = pd.concat(load_parts, ignore_index=True).groupby(["PLANT", "RESOURCE"], dropna=False, observed=True, as_index=False)["OUTPUT_LOAD_MT"].sum()
        out = out.merge(load, on=["PLANT", "RESOURCE"], how="left")
        out["LOAD_MT"] = out["LOAD_MT"].fillna(out.pop("OUTPUT_LOAD_MT"))
    return out.sort_values(["BOTTLENECK", "UTILIZATION_PCT"], ascending=[False, False], na_position="last").reset_index(drop=True)


def build_network_trace(sources: Mapping[str, SourceTable]) -> pd.DataFrame:
    rows: list[pd.DataFrame] = []
    for source in sources.values():
        df, n = source.data, source.name
        if df.empty: continue
        if "ESP_BOM" in n:
            parent = first_column(df, ("PARENTITEM", "PARENT", "OUTPUTITEM", "ASSEMBLY", "ITEM"))
            component = first_column(df, ("COMPONENTITEM", "COMPONENT", "INPUTITEM", "SUBITEM"))
            loc = canonical_column(df, "PLANT")
            if parent and component:
                rows.append(pd.DataFrame({"EDGE_TYPE": "BOM", "FROM_NODE": clean_text(df[component]), "TO_NODE": clean_text(df[parent]),
                                          "ITEM": clean_text(df[parent]), "PRODUCING_PLANT": clean_text(df[loc]) if loc else "(Blank)", "RESOURCE": "(Blank)", "SOURCE_FILE": source.path.name}))
        if "ESP_PRODUCTIONSTEP" in n or "ESP_PRODUCTIONMETHOD" in n:
            item = canonical_column(df, "ITEM"); loc = canonical_column(df, "PLANT"); res = canonical_column(df, "RESOURCE")
            method = first_column(df, ("PRODUCTIONMETHOD", "METHOD", "PRODMETHOD", "STEP"))
            if item:
                rows.append(pd.DataFrame({"EDGE_TYPE": "Production route", "FROM_NODE": clean_text(df[method]) if method else "Production",
                                          "TO_NODE": clean_text(df[item]), "ITEM": clean_text(df[item]),
                                          "PRODUCING_PLANT": clean_text(df[loc]) if loc else "(Blank)", "RESOURCE": clean_text(df[res]) if res else "(Blank)", "SOURCE_FILE": source.path.name}))
        if "ESP_PLANARRIV" in n:
            item = canonical_column(df, "ITEM"); src = first_column(df, ("SOURCE", "ORIGIN")); dest = first_column(df, ("DEST", "DESTINATION", "LOC"))
            if src and dest:
                rows.append(pd.DataFrame({"EDGE_TYPE": "Plan arrival", "FROM_NODE": clean_text(df[src]), "TO_NODE": clean_text(df[dest]),
                                          "ITEM": clean_text(df[item]) if item else "(Blank)", "PRODUCING_PLANT": clean_text(df[src]),
                                          "RESOURCE": "(Blank)", "SOURCE_FILE": source.path.name}))
    columns = ["EDGE_TYPE", "FROM_NODE", "TO_NODE", "ITEM", "PRODUCING_PLANT", "RESOURCE", "SOURCE_FILE"]
    return pd.concat(rows, ignore_index=True)[columns].drop_duplicates().reset_index(drop=True) if rows else pd.DataFrame(columns=columns)


def build_network_lanes(sources: Mapping[str, SourceTable]) -> pd.DataFrame:
    """Weekly ESP_PLANARRIV lanes used by the app/Excel network animation."""
    parts = []
    preferred = _latest_matching_source(sources, ("ESP_PLANARRIV",))
    if preferred is None or preferred.data.empty:
        return pd.DataFrame(columns=["WEEK", "SOURCE", "DEST", "ITEM", "QTY_MT"])
    df = preferred.data
    src = first_column(df, ("SOURCE", "ORIGIN")); dst = first_column(df, ("DEST", "DESTINATION"))
    item = canonical_column(df, "ITEM"); qty = first_column(df, ("QTY", "PLANQTY"))
    date = first_column(df, ("SCHEDARRIVDATE", "SCHEDSHIPDATE", "NEEDDATE"))
    if not src or not dst or not qty or not date:
        return pd.DataFrame(columns=["WEEK", "SOURCE", "DEST", "ITEM", "QTY_MT"])
    lanes = pd.DataFrame({
        "WEEK": week_start(parse_dates(df[date])), "SOURCE": clean_text(df[src]), "DEST": clean_text(df[dst]),
        "ITEM": clean_text(df[item]) if item else "(Blank)", "QTY_MT": quantity_mt(df[qty], qty),
    })
    start = _inventory_plan_start(sources)
    end = _resource_plan_end(sources)
    if start is not None:
        lanes = lanes[lanes["WEEK"].ge(start)]
    if end is not None:
        lanes = lanes[lanes["WEEK"].le(end)]
    return (lanes.groupby(["WEEK", "SOURCE", "DEST", "ITEM"], dropna=False, observed=True, as_index=False)["QTY_MT"].sum()
            .sort_values(["WEEK", "QTY_MT"], ascending=[True, False]).reset_index(drop=True))


def network_animation_gif(model: TwinModel, frame_phases: int = 2) -> bytes:
    """Create the weekly network architecture clip as an animated GIF."""
    from PIL import Image, ImageDraw, ImageFont

    width, height, panel = 1080, 610, 310
    bg, grid, cyan, yellow = "#0B1724", "#183147", "#36C4E7", "#FFE36E"
    positions = {
        "SIL": (110, 275), "TLJ": (115, 445), "MOU": (280, 370), "MHM": (380, 245),
        "RKT": (495, 115), "HIR": (455, 455), "AFG": (565, 335), "BEL": (655, 235),
        "V_WEST": (45, 150), "V_NORTH": (305, 60), "V_EAST": (655, 95),
        "V_EXPORT": (75, 555), "V_SOUTH": (620, 555),
    }
    lanes = model.tables.get("network_lanes_weekly", pd.DataFrame()).copy()
    trend = model.tables.get("weekly_trend", pd.DataFrame()).copy()
    balance = model.tables.get("inventory_balance_wide", pd.DataFrame()).copy()
    capacity = model.tables.get("capacity", pd.DataFrame()).copy()
    weeks = sorted(set(pd.to_datetime(lanes.get("WEEK", pd.Series(dtype="datetime64[ns]")), errors="coerce").dropna()) |
                   set(pd.to_datetime(trend.get("WEEK", pd.Series(dtype="datetime64[ns]")), errors="coerce").dropna()))
    if not weeks:
        weeks = [pd.Timestamp(model.generated_at).normalize()]

    # Resource stage lookup and static bottleneck labels.
    rp = model.tables.get("resource_prod_weekly", pd.DataFrame())
    stage_map = {}
    if not rp.empty and {"RESOURCE", "PRODUCT_FORM"}.issubset(rp.columns):
        for key, grp in rp.groupby(["PLANT", "RESOURCE"], dropna=False, observed=True):
            stage_map[key] = str(grp.sort_values("QTY_MT", ascending=False).iloc[0].get("PRODUCT_FORM", ""))
    bott = capacity[capacity.get("BOTTLENECK", pd.Series(False, index=capacity.index)).fillna(False)].copy() if not capacity.empty else capacity
    bott_by_plant = bott.groupby("PLANT", dropna=False, observed=True).size().to_dict() if not bott.empty else {}

    week_cols = {pd.to_datetime(c, format="%d-%b-%y", errors="coerce"): c for c in balance.columns if isinstance(c, str)}
    def stock(week: pd.Timestamp, category: str) -> tuple[float, dict[str, float]]:
        col = next((name for dtv, name in week_cols.items() if pd.notna(dtv) and dtv.normalize() == week.normalize()), None)
        if col is None or balance.empty:
            return 0.0, {}
        x = balance[(balance["Measure"] == "Closing inventory") & (balance["Category"] == category)]
        by = x.groupby("Plant", observed=True)[col].sum().to_dict()
        return float(sum(by.values())), {str(k): float(v) for k, v in by.items()}

    font = ImageFont.load_default(); bold = ImageFont.load_default()
    frames = []
    for week in weeks:
        wk_lanes = lanes[pd.to_datetime(lanes.get("WEEK"), errors="coerce").eq(week)] if not lanes.empty else lanes
        active = wk_lanes.groupby(["SOURCE", "DEST"], as_index=False)["QTY_MT"].sum().nlargest(18, "QTY_MT") if not wk_lanes.empty else pd.DataFrame()
        fg_total, fg_plant = stock(week, "FG"); wip_total, wip_plant = stock(week, "WIP")
        prod = float(trend[(trend.get("WEEK") == week) & trend.get("MEASURE", pd.Series(index=trend.index)).eq("Production")]["QTY_MT"].sum()) if not trend.empty else 0.0
        cust = float(wk_lanes[wk_lanes.get("DEST", pd.Series(dtype=str)).astype(str).str.startswith("V_")]["QTY_MT"].sum()) if not wk_lanes.empty else 0.0
        for phase in range(max(1, frame_phases)):
            im = Image.new("RGB", (width, height), bg); d = ImageDraw.Draw(im)
            for x in range(0, width-panel, 40): d.line((x, 0, x, height), fill=grid, width=1)
            for y in range(0, height, 40): d.line((0, y, width-panel, y), fill=grid, width=1)
            d.text((18, 12), "Hindalco Supply Network Digital Twin | Blue Yonder ESP run", fill="white", font=bold)
            d.text((18, 30), "Weekly material flow (MT), plant & machine-centre status — schematic, not to scale", fill="#A8B8C8", font=font)
            # Base and active lanes.
            lane_pairs = [(str(r.SOURCE), str(r.DEST), float(r.QTY_MT)) for r in active.itertuples(index=False)] if not active.empty else []
            for src, dst, qty in lane_pairs:
                if src not in positions or dst not in positions: continue
                a, b = positions[src], positions[dst]
                d.line((*a, *b), fill=cyan, width=2)
                t = (phase + 1) / (max(1, frame_phases) + 1)
                mx, my = a[0] + (b[0]-a[0])*t, a[1] + (b[1]-a[1])*t
                d.rounded_rectangle((mx-7, my-4, mx+7, my+4), radius=2, fill=yellow)
            # Nodes.
            for node, (x, y) in positions.items():
                if node.startswith("V_"):
                    d.rounded_rectangle((x-34, y-11, x+34, y+11), radius=5, fill="#65428A", outline="#A884D3")
                    d.text((x-27, y-6), node, fill="white", font=font)
                else:
                    d.ellipse((x-22, y-22, x+22, y+22), fill="#4B5968", outline="#F05A34", width=3)
                    d.text((x-11, y-6), node, fill="white", font=bold)
                    d.text((x-31, y+27), f"FG {fg_plant.get(node,0)/1000:.1f}k | WIP {wip_plant.get(node,0)/1000:.1f}k", fill="#AFC4D4", font=font)
                    if bott_by_plant.get(node, 0): d.text((x+20, y-30), f"{bott_by_plant[node]}x100%", fill="#FF4438", font=font)
            # Side panel.
            px = width-panel; d.rectangle((px, 0, width, height), fill="#15283A")
            d.rectangle((px, 0, px+5, height), fill="#F47732")
            d.text((px+18, 20), "Week bucket", fill="#AFC4D4", font=font); d.text((px+18, 42), week.strftime("%d-%b-%y"), fill="white", font=bold)
            metrics = [("FG production", prod), ("Customer dispatch", cust), ("Network FG closing", fg_total), ("Network WIP closing", wip_total)]
            y = 78
            for label, value in metrics:
                d.rounded_rectangle((px+14, y, width-14, y+48), radius=6, fill="#20374B")
                d.text((px+23, y+8), label, fill="#9DB2C4", font=font); d.text((px+23, y+26), f"{value:,.0f} MT", fill="white", font=bold); y += 57
            d.text((px+18, y+2), f"Bottleneck machines: {len(bott)}", fill="#FF8A45", font=bold); y += 24
            top = bott.sort_values(["UTILIZATION_PCT", "RESOURCE"], ascending=[False, True]).head(9) if not bott.empty else bott
            for r in top.itertuples(index=False):
                stage = stage_map.get((getattr(r, "PLANT", ""), getattr(r, "RESOURCE", "")), "")
                text = f"{getattr(r,'PLANT','')}  {getattr(r,'RESOURCE','')} ({stage})"
                d.text((px+18, y), text[:45], fill="#D5E0E8", font=font); y += 17
            frames.append(im)
    out = io.BytesIO()
    frames[0].save(out, format="GIF", save_all=True, append_images=frames[1:], duration=500, loop=0, optimize=True)
    return out.getvalue()


def build_root_cause_misses(service: pd.DataFrame, trace: pd.DataFrame) -> pd.DataFrame:
    if service.empty:
        return pd.DataFrame(columns=["PRIORITY", "BU", "DEMAND_LOCATION", "ITEM", "UNALLOCATED_MT", "PRODUCING_PLANT", "RESOURCE", "TRACE_STATUS"])
    misses = (service.groupby(["BU", "PLANT", "ITEM"], dropna=False, observed=True, as_index=False)["UNALLOCATED_MT"].sum()
              .query("UNALLOCATED_MT > 0").rename(columns={"PLANT": "DEMAND_LOCATION"}))
    routes = trace[trace["EDGE_TYPE"].isin(["Plan arrival", "Production route"])].copy()
    plan = routes[routes["EDGE_TYPE"].eq("Plan arrival")].rename(columns={"TO_NODE": "DEMAND_LOCATION"})
    plan = plan[["DEMAND_LOCATION", "ITEM", "PRODUCING_PLANT"]].drop_duplicates()
    out = misses.merge(plan, on=["DEMAND_LOCATION", "ITEM"], how="left")
    prod = routes[routes["EDGE_TYPE"].eq("Production route")][["ITEM", "PRODUCING_PLANT", "RESOURCE"]].drop_duplicates()
    out = out.merge(prod, on=["ITEM", "PRODUCING_PLANT"], how="left")
    out["TRACE_STATUS"] = np.where(out["PRODUCING_PLANT"].isna(), "No producing route found", np.where(out["RESOURCE"].isna(), "Plant traced; no resource", "Plant and resource traced"))
    out["PRIORITY"] = np.where(out["UNALLOCATED_MT"].gt(out["UNALLOCATED_MT"].quantile(.75)), "P1", "P2")
    return out[["PRIORITY", "BU", "DEMAND_LOCATION", "ITEM", "UNALLOCATED_MT", "PRODUCING_PLANT", "RESOURCE", "TRACE_STATUS"]].sort_values("UNALLOCATED_MT", ascending=False).reset_index(drop=True)


def build_order_status(sources: Mapping[str, SourceTable]) -> pd.DataFrame:
    frames = []
    plan_start, plan_end, _ = _plan_horizon(sources)
    inventory_start = _inventory_plan_start(sources)
    if inventory_start is not None:
        plan_start = inventory_start
    preferred = _latest_matching_source(sources, ("ESP_CUSTORDERSTATUS", "CUSTORDERSTATUS"))
    for source in sources.values():
        if "ESP_CUSTORDERSTATUS" not in source.name and "CUSTORDERSTATUS" not in source.name: continue
        if preferred is not None and source is not preferred: continue
        df = source.data
        order_type_col = first_column(df, ("U_ORDER_TYPE", "ORDER_TYPE", "TYPE"))
        if order_type_col:
            types = clean_text(df[order_type_col], "").str.upper()
            cust = types.str.contains("CUST|CUSTOMER", regex=True)
            if cust.any(): df = df.loc[cust]
        due = first_column(df, ("SCHEDDATE", "DUEDATE", "NEEDDATE", "REQUESTDATE", "ORDERDATE"))
        if not due: continue
        dims = dimension_frame(df); due_dates = parse_dates(df[due])
        order_col = first_column(df, ("ORDERQTY", "QTY", "DEMANDQTY")); delivered_col = first_column(df, ("DELIVEREDQTY", "SHIPQTY", "FULFILLEDQTY", "SCHEDQTY"))
        delivery_date_col = first_column(df, ("SHIPDATE", "DELIVERYDATE", "ACTUALDELIVERYDATE"))
        otif_col = first_column(df, ("OTIF", "IS_OTIF", "ONTIMEINFULL")); ontime_col = first_column(df, ("ONTIME", "ON_TIME")); infull_col = first_column(df, ("INFULL", "IN_FULL"))
        out = dims.copy(); out["DUE_DATE"] = due_dates; out["ORDER_MT"] = quantity_mt(df[order_col], order_col) if order_col else 0.0
        out["DELIVERED_MT"] = quantity_mt(df[delivered_col], delivered_col) if delivered_col else 0.0
        out["DELIVERY_DATE"] = parse_dates(df[delivery_date_col]) if delivery_date_col else pd.NaT
        if otif_col:
            text = clean_text(df[otif_col], "").str.upper(); out["OTIF"] = text.isin(["1", "Y", "YES", "TRUE", "OTIF", "ON TIME IN FULL"])
        elif ontime_col and infull_col:
            yes = lambda s: clean_text(s, "").str.upper().isin(["1", "Y", "YES", "TRUE"])
            out["OTIF"] = yes(df[ontime_col]) & yes(df[infull_col])
        else:
            status_col = first_column(df, ("SCHEDSTATUS", "SCHEDULESTATUS", "STATUS"))
            if status_col:
                status = clean_text(df[status_col], "").str.upper()
                # Real ESP values are prefixed, e.g. "1 - Met", "3 - Met late"
                # and "5 - Not met".  The supplied definition accepts a met or
                # completed line but explicitly rejects NOT MET.
                met = ((status.str.contains(r"\bMET\b|FULFILLED|COMPLETE", regex=True))
                       & ~status.str.contains(r"NOT\s+MET", regex=True))
            else:
                met = pd.Series(False, index=df.index)
            # Business definition supplied for this ESP status output.
            out["OTIF"] = met & out["DUE_DATE"].le(out["DELIVERY_DATE"])
        out["BEFORE_PLAN_START"] = out["DELIVERY_DATE"].lt(plan_start) if plan_start is not None else False
        out["ELIGIBLE_IN_HORIZON"] = ~out["BEFORE_PLAN_START"] & out["DELIVERY_DATE"].notna() & out["DUE_DATE"].notna()
        out["OTIF_WEIGHT_MT"] = np.where(out["OTIF"] & out["ELIGIBLE_IN_HORIZON"], out["ORDER_MT"], 0.0)
        out["PAST_DUE_EXCLUDED"] = out["BEFORE_PLAN_START"]
        out["MATCHED_QTY_COLUMN"] = order_col or ""; out["MATCHED_SCHEDQTY_COLUMN"] = delivered_col or ""
        out["MATCHED_DUE_COLUMN"] = due; out["MATCHED_SHIP_COLUMN"] = delivery_date_col or ""
        out["SOURCE_FILE"] = source.path.name
        frames.append(out)
    columns = ["DUE_DATE", "DELIVERY_DATE", "BU", "PLANT", "LOCATION_TYPE", "ITEM", "ORDER_MT", "DELIVERED_MT", "OTIF", "OTIF_WEIGHT_MT",
               "BEFORE_PLAN_START", "ELIGIBLE_IN_HORIZON", "PAST_DUE_EXCLUDED",
               "MATCHED_QTY_COLUMN", "MATCHED_SCHEDQTY_COLUMN", "MATCHED_DUE_COLUMN", "MATCHED_SHIP_COLUMN", "SOURCE_FILE"]
    if not frames: return pd.DataFrame(columns=columns)
    return pd.concat(frames, ignore_index=True)[columns].sort_values("DUE_DATE").reset_index(drop=True)


def build_inventory_health(sources: Mapping[str, SourceTable]) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    inventory_rows = []
    placeholders = []
    overdue = []
    today = pd.Timestamp.now().normalize()
    for source in sources.values():
        df, n = source.data, source.name
        if n == "ESP_INVENTORY" or ("ESP_INVENTORY" in n and "PARAM" not in n):
            item = canonical_column(df, "ITEM"); loc = canonical_column(df, "PLANT"); date = first_column(df, ("AVAILDATE", "DATE")); qty = first_column(df, ("PROJECTED_ON_HAND", "QTY", "ONHANDQTY"))
            if item and date and qty:
                x = pd.DataFrame({"ITEM": clean_text(df[item]), "PLANT": clean_text(df[loc]) if loc else "(Blank)", "DATE": parse_dates(df[date]), "PROJECTED_ON_HAND_MT": quantity_mt(df[qty], qty)})
                for (it, pl), g in x.dropna(subset=["DATE"]).sort_values("DATE").groupby(["ITEM", "PLANT"], dropna=False, observed=True):
                    first, last = float(g["PROJECTED_ON_HAND_MT"].iloc[0]), float(g["PROJECTED_ON_HAND_MT"].iloc[-1])
                    inventory_rows.append({"ITEM": it, "PLANT": pl, "FIRST_DATE": g["DATE"].iloc[0], "LAST_DATE": g["DATE"].iloc[-1], "OPENING_MT": first, "CLOSING_MT": last,
                                           "CHANGE_MT": last-first, "NON_MOVING": len(g)>1 and g["PROJECTED_ON_HAND_MT"].nunique(dropna=False)==1 and last>0, "BUILD_UP": last>first})
        if "ESP_SKUEFFINVENTORYPARAM" in n:
            dims = dimension_frame(df)
            for col in df.columns:
                vals = pd.to_numeric(df[col], errors="coerce")
                mask = vals.abs().ge(999_999)
                if mask.any():
                    p = dims.loc[mask, ["BU", "PLANT", "ITEM"]].copy(); p["PARAMETER"] = col; p["VALUE"] = vals.loc[mask]; p["ISSUE"] = "999,999 placeholder"
                    placeholders.append(p)
        if "ESP_PURCHORDER" in n:
            due_col = first_column(df, ("DUEDATE", "NEEDDATE")); qty_col = first_column(df, ("REMAININGQTY", "OPENQTY", "QTY"))
            if due_col and qty_col:
                dims = dimension_frame(df); due_dates = parse_dates(df[due_col]); open_mt = quantity_mt(df[qty_col], qty_col)
                mask = due_dates.lt(today) & open_mt.gt(0)
                if mask.any():
                    p = dims.loc[mask, ["BU", "PLANT", "ITEM"]].copy(); p["DUE_DATE"] = due_dates.loc[mask]; p["OPEN_MT"] = open_mt.loc[mask]; p["DAYS_OVERDUE"] = (today-due_dates.loc[mask]).dt.days
                    overdue.append(p)
    inv_cols = ["ITEM", "PLANT", "FIRST_DATE", "LAST_DATE", "OPENING_MT", "CLOSING_MT", "CHANGE_MT", "NON_MOVING", "BUILD_UP"]
    ph_cols = ["BU", "PLANT", "ITEM", "PARAMETER", "VALUE", "ISSUE"]
    po_cols = ["BU", "PLANT", "ITEM", "DUE_DATE", "OPEN_MT", "DAYS_OVERDUE"]
    return (pd.DataFrame(inventory_rows, columns=inv_cols).sort_values("CHANGE_MT", ascending=False),
            pd.concat(placeholders, ignore_index=True)[ph_cols] if placeholders else pd.DataFrame(columns=ph_cols),
            pd.concat(overdue, ignore_index=True)[po_cols].sort_values("OPEN_MT", ascending=False) if overdue else pd.DataFrame(columns=po_cols))


def build_sourcing_checks(sources: Mapping[str, SourceTable]) -> tuple[pd.DataFrame, pd.DataFrame]:
    facts, issues = [], []
    for source in sources.values():
        if "ESP_SOURCING" not in source.name: continue
        df = source.data; item = canonical_column(df, "ITEM"); src = first_column(df, ("SOURCE", "ORIGIN", "SOURCELOC")); dest = first_column(df, ("DEST", "DESTINATION", "LOC", "LOCATION")); factor = first_column(df, ("FACTOR", "SOURCINGFACTOR", "PERCENT", "ALLOCATION"))
        if not item: continue
        x = pd.DataFrame({"ITEM": clean_text(df[item]), "SOURCE": clean_text(df[src], "") if src else "", "DEST": clean_text(df[dest], "") if dest else "", "FACTOR": numeric(df[factor]) if factor else np.nan})
        x["SOURCE_FILE"] = source.path.name; facts.append(x)
        missing = x["SOURCE"].eq("") | x["DEST"].eq("")
        for _, r in x[missing].iterrows(): issues.append({"SEVERITY": "P1", "ITEM": r.ITEM, "SOURCE": r.SOURCE, "DEST": r.DEST, "ISSUE": "Missing source or destination"})
        if factor:
            bad = x["FACTOR"].le(0)
            for _, r in x[bad].iterrows(): issues.append({"SEVERITY": "P1", "ITEM": r.ITEM, "SOURCE": r.SOURCE, "DEST": r.DEST, "ISSUE": "Invalid sourcing factor"})
        dup = x.duplicated(["ITEM", "SOURCE", "DEST"], keep=False)
        for _, r in x[dup].drop_duplicates(["ITEM", "SOURCE", "DEST"]).iterrows(): issues.append({"SEVERITY": "P2", "ITEM": r.ITEM, "SOURCE": r.SOURCE, "DEST": r.DEST, "ISSUE": "Duplicate sourcing lane"})
    fact_cols = ["ITEM", "SOURCE", "DEST", "FACTOR", "SOURCE_FILE"]; issue_cols = ["SEVERITY", "ITEM", "SOURCE", "DEST", "ISSUE"]
    return (pd.concat(facts, ignore_index=True)[fact_cols] if facts else pd.DataFrame(columns=fact_cols), pd.DataFrame(issues, columns=issue_cols))


def build_master_health(sources: Mapping[str, SourceTable], item_master: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for source in sources.values():
        if not any(token in source.name for token in ("ESP_SKU", "ESP_SOURCING", "ESP_BOM", "ESP_PRODUCTIONSTEP", "ESP_PRODUCTIONMETHOD", "DFUTOSKU_FE")):
            continue
        df = source.data; total = max(len(df), 1)
        for field in ("ITEM", "PLANT"):
            col = canonical_column(df, field)
            if not col:
                rows.append({"STATUS": "FAIL", "SEVERITY": "P1", "FILE": source.path.name, "CHECK": f"{field} column present", "MATCHED_COLUMN": "", "COUNT": len(df), "RATE_PCT": 100.0})
            else:
                missing = clean_text(df[col], "").eq("").sum()
                rows.append({"STATUS": "PASS" if not missing else "FAIL", "SEVERITY": "Info" if not missing else "P1", "FILE": source.path.name,
                             "CHECK": f"Blank {field}", "MATCHED_COLUMN": col, "COUNT": int(missing), "RATE_PCT": missing/total*100})
        dup = df.duplicated().sum()
        rows.append({"STATUS": "PASS" if not dup else "WARN", "SEVERITY": "Info" if not dup else "P2", "FILE": source.path.name,
                     "CHECK": "Duplicate rows", "MATCHED_COLUMN": "All columns", "COUNT": int(dup), "RATE_PCT": dup/total*100})
    if not item_master.empty:
        blank = item_master["BU"].eq("(Blank)").sum()
        rows.append({"STATUS": "PASS" if not blank else "FAIL", "SEVERITY": "Info" if not blank else "P1", "FILE": "Item master crosswalk",
                     "CHECK": "Items without BU", "MATCHED_COLUMN": "BU", "COUNT": int(blank), "RATE_PCT": blank/len(item_master)*100})
    columns = ["STATUS", "SEVERITY", "FILE", "CHECK", "MATCHED_COLUMN", "COUNT", "RATE_PCT"]
    return pd.DataFrame(rows, columns=columns).sort_values(["STATUS", "COUNT"], ascending=[True, False]).reset_index(drop=True)


def build_diagnostics(sources: Mapping[str, SourceTable]) -> pd.DataFrame:
    rows = []
    for source in sources.values():
        df = source.data
        matched = {field: canonical_column(df, field) or "" for field in ("DATE", "BU", "PLANT", "RESOURCE", "ITEM", "DFU", "DMDGROUP", "QTY")}
        step_errors = []
        extra_matches: list[str] = []
        n = source.name
        if "UDT_ORDER_FULFILLMENT" in n:
            for label, aliases in {
                "TOTAL_ORDER": ("TOTAL_ORDER",), "TOTAL_ALLOCATION": ("TOTAL_ALLOCATION",), "MONTH": ("MONTH",),
            }.items():
                found = first_column(df, aliases)
                if not found: step_errors.append(f"Missing {label}")
                else: extra_matches.append(f"{label}={found}")
            bu_col = canonical_column(df, "BU")
            if bu_col:
                blank_bu = int(clean_text(df[bu_col], "").eq("").sum())
                extra_matches.append(f"BLANK_BU_ROWS_EXCLUDED={blank_bu}")
        if "RESOURCE_UTILIZATION" in n:
            date_headers = sum(pd.notna(parse_dates(pd.Series([c])).iloc[0]) for c in df.columns)
            if not canonical_column(df, "RESOURCE"): step_errors.append("Missing RESOURCE")
            if not date_headers: step_errors.append("No weekly date headers matched")
            extra_matches.append(f"WEEKLY_DATE_HEADERS={date_headers}")
        if "ESP_CUSTORDERSTATUS" in n:
            for label, aliases in {"QTY": ("QTY",), "SCHEDQTY": ("SCHEDQTY",), "SHIPDATE": ("SHIPDATE",), "SCHEDDATE": ("SCHEDDATE",)}.items():
                found = first_column(df, aliases)
                if not found: step_errors.append(f"Missing {label}")
                else: extra_matches.append(f"{label}={found}")
        rows.append({
            "FILE": source.path.name, "ROWS_BEFORE_SIMULATION": source.rows_before_filter,
            "ROWS_AFTER_SIMULATION": source.rows_after_filter, "ROWS_REMOVED": source.rows_before_filter-source.rows_after_filter,
            "SIMULATION_FILTER": source.simulation_note, **{f"MATCHED_{k}": v for k, v in matched.items()},
            "MATCHED_EXTRA": "; ".join(extra_matches),
            "STEP_ERRORS": source.load_error or "; ".join(step_errors),
        })
    return pd.DataFrame(rows).sort_values(["STEP_ERRORS", "FILE"], ascending=[False, True]).reset_index(drop=True)


def build_where_to_hit(service: pd.DataFrame, capacity: pd.DataFrame, forecast_loss: pd.DataFrame, margin: pd.DataFrame, overdue: pd.DataFrame,
                       sourcing_issues: pd.DataFrame, placeholders: pd.DataFrame, inventory_health: pd.DataFrame,
                       master_health: pd.DataFrame) -> pd.DataFrame:
    rows = []
    if not service.empty:
        for _, r in service.groupby(["BU", "PLANT", "ITEM"], dropna=False, observed=True, as_index=False)["UNALLOCATED_MT"].sum().nlargest(100, "UNALLOCATED_MT").iterrows():
            if r.UNALLOCATED_MT > 0: rows.append({"PRIORITY": "P1", "AREA_RANK": 0, "AREA": "Service", "WHERE": f"{r.BU} / {r.PLANT} / {r.ITEM}", "IMPACT_MT": r.UNALLOCATED_MT, "IMPACT": f"{r.UNALLOCATED_MT:,.1f} MT unallocated", "ACTION": "Trace producing plant/resource and recover constrained supply."})
    if not capacity.empty:
        for _, r in capacity[capacity["BOTTLENECK"]].head(100).iterrows(): rows.append({"PRIORITY": "P1", "AREA_RANK": 2, "AREA": "Capacity", "WHERE": f"{r.PLANT} / {r.RESOURCE}", "IMPACT_MT": 0.0, "IMPACT": f"{r.UTILIZATION_PCT:,.1f}% peak utilization", "ACTION": "Validate weekly load, shutdown calendar, alternate resource and timing."})
    if not forecast_loss.empty:
        for _, r in forecast_loss.nlargest(100, "LOST_MT").iterrows():
            rows.append({"PRIORITY": "P1", "AREA_RANK": 1, "AREA": "Forecast mapping", "WHERE": f"{r.BU} / {r.LOC} / {r.DFU}",
                         "IMPACT_MT": r.LOST_MT, "IMPACT": f"{r.LOST_MT:,.1f} MT lost — {r.REASON}", "ACTION": "Correct DFUTOSKU_FE or downstream SKU forecast and rerun."})
    if not margin.empty:
        scheduled = float(margin["NEGATIVE_MARGIN_SCHEDULED_MT"].sum()); value = float(margin["VALUE_AT_RISK"].sum())
        if scheduled > 0 or value > 0:
            rows.append({"PRIORITY": "P2", "AREA_RANK": 6, "AREA": "Margin", "WHERE": "All negative-margin demand",
                         "IMPACT_MT": scheduled, "IMPACT": f"{scheduled:,.1f} MT negative-margin scheduled", "ACTION": "Review the aggregated margin detail by BU/plant/item before release."})
    for _, r in overdue.head(100).iterrows(): rows.append({"PRIORITY": "P1", "AREA_RANK": 3, "AREA": "Purchase", "WHERE": f"{r.PLANT} / {r.ITEM}", "IMPACT_MT": r.OPEN_MT, "IMPACT": f"{r.OPEN_MT:,.1f} MT, {int(r.DAYS_OVERDUE)} days overdue", "ACTION": "Expedite, reschedule or close the purchase order."})
    for _, r in sourcing_issues.head(100).iterrows(): rows.append({"PRIORITY": r.SEVERITY, "AREA_RANK": 4, "AREA": "Sourcing", "WHERE": f"{r.SOURCE} → {r.DEST} / {r.ITEM}", "IMPACT_MT": 0.0, "IMPACT": r.ISSUE, "ACTION": "Correct sourcing master and regenerate the plan."})
    for _, r in placeholders.head(100).iterrows(): rows.append({"PRIORITY": "P1", "AREA_RANK": 5, "AREA": "Inventory parameter", "WHERE": f"{r.PLANT} / {r.ITEM}", "IMPACT_MT": 0.0, "IMPACT": f"{r.PARAMETER}={r.VALUE:,.0f}", "ACTION": "Replace 999,999 placeholder with an approved parameter."})
    for _, r in inventory_health[(inventory_health.get("NON_MOVING", False)==True) | (inventory_health.get("BUILD_UP", False)==True)].head(100).iterrows(): rows.append({"PRIORITY": "P2", "AREA_RANK": 7, "AREA": "Inventory", "WHERE": f"{r.PLANT} / {r.ITEM}", "IMPACT_MT": abs(r.CHANGE_MT), "IMPACT": f"Change {r.CHANGE_MT:,.1f} MT; non-moving={r.NON_MOVING}", "ACTION": "Review projected-on-hand profile, demand and receipts."})
    failures = master_health[master_health.get("STATUS", pd.Series(dtype=str)).ne("PASS")]
    for _, r in failures.head(100).iterrows(): rows.append({"PRIORITY": r.SEVERITY, "AREA_RANK": 8, "AREA": "Master data", "WHERE": r.FILE, "IMPACT_MT": 0.0, "IMPACT": f"{r.CHECK}: {int(r.COUNT):,}", "ACTION": "Correct the source master and re-export."})
    cols = ["PRIORITY", "AREA", "WHERE", "IMPACT_MT", "IMPACT", "ACTION"]
    if not rows: return pd.DataFrame(columns=cols)
    out = pd.DataFrame(rows).sort_values(["PRIORITY", "AREA_RANK", "IMPACT_MT"], ascending=[True, True, False]).drop(columns="AREA_RANK")
    return out[cols].reset_index(drop=True)


def build_twin(folder: str | Path, simulation: str | None = None) -> TwinModel:
    sources, warnings = load_sources(folder, simulation)
    model = TwinModel(Path(folder), simulation, sources=sources, warnings=warnings)
    item_master = build_item_master(sources)
    trend = build_weekly_trend(sources)
    resource_prod = build_resource_production(sources, item_master)
    if resource_prod.empty:
        resource_prod = aggregate_trend(trend, "WEEK", ["BU", "PLANT", "RESOURCE", "PRODUCT_FORM", "STREAM", "SUB_CLASS"], ["Production"])
    bu_head = build_bu_head_forecast(sources)
    flow = enrich_with_item_master(build_forecast_flow(sources), item_master)
    loss = build_dfu_mapping_loss(sources)
    capacity = build_capacity(sources)
    bottlenecks = capacity if not capacity.empty else build_bottlenecks(sources, trend)
    order_priority = build_order_priority(sources)
    service = enrich_with_item_master(build_service_demand(sources), item_master)
    margin = enrich_with_item_master(build_value_at_risk(sources), item_master)
    network = build_network_trace(sources)
    root_causes = build_root_cause_misses(service, network)
    order_status = enrich_with_item_master(build_order_status(sources), item_master)
    inventory_health, inventory_placeholders, overdue_po = build_inventory_health(sources)
    sourcing, sourcing_issues = build_sourcing_checks(sources)
    master_health = build_master_health(sources, item_master)
    network_lanes = build_network_lanes(sources)
    inventory_balance_wide = pd.DataFrame()
    inventory_balance_weeks: list[str] = []
    try:
        # The standalone report and the main app share one implementation so
        # balances, hierarchy and Excel formulas cannot drift apart.
        import inv_balance_report as _ibr
        inventory_balance_wide, inventory_balance_weeks = _ibr.build(str(folder))
    except Exception as exc:
        warnings.append(f"Inventory drill-down could not be built: {exc}")
    where_to_hit = build_where_to_hit(service, bottlenecks, loss, margin, overdue_po, sourcing_issues,
                                      inventory_placeholders, inventory_health, master_health)
    forecast_value_cols = [c for c in ["BU_FCST_MT", "DFUTOSKU_FCST_MT", "NETTED_FCST_MT", "CUSTORDER_MT", "TOTAL_FCST_MT", "SCHED_QTY_MT"] if c in flow]
    forecast_summary = (flow.groupby(["MONTH", "BU", "PLANT"], dropna=False, observed=True, as_index=False)[forecast_value_cols].sum()
                        if not flow.empty else pd.DataFrame(columns=["MONTH", "BU", "PLANT", *forecast_value_cols]))
    if not forecast_summary.empty:
        forecast_summary["MET_PCT"] = np.where(forecast_summary["TOTAL_FCST_MT"].gt(0), forecast_summary["SCHED_QTY_MT"] / forecast_summary["TOTAL_FCST_MT"] * 100, np.nan)
    service_summary = (service.groupby(["MONTH" if "MONTH" in service else "DATE", "BU", "PLANT"], dropna=False, observed=True, as_index=False)[["DEMAND_MT", "ALLOCATED_MT", "UNALLOCATED_MT"]].sum()
                       if not service.empty else pd.DataFrame(columns=["DATE", "BU", "PLANT", "DEMAND_MT", "ALLOCATED_MT", "UNALLOCATED_MT"]))
    if not service_summary.empty:
        service_summary["ORDER_FILL_PCT"] = np.where(service_summary["DEMAND_MT"].gt(0), service_summary["ALLOCATED_MT"] / service_summary["DEMAND_MT"] * 100, np.nan)
    otif_summary = (order_status.groupby(["BU", "PLANT"], dropna=False, observed=True, as_index=False)[["ORDER_MT", "OTIF_WEIGHT_MT"]].sum()
                    if not order_status.empty else pd.DataFrame(columns=["BU", "PLANT", "ORDER_MT", "OTIF_WEIGHT_MT"]))
    if not otif_summary.empty:
        otif_summary["OTIF_PCT"] = np.where(otif_summary["ORDER_MT"].gt(0), otif_summary["OTIF_WEIGHT_MT"] / otif_summary["ORDER_MT"] * 100, np.nan)

    model.tables.update({
        "weekly_trend": trend,
        "resource_prod_weekly": resource_prod,
        "resource_prod_monthly": (resource_prod.groupby(["MONTH", "BU", "PLANT", "RESOURCE", "PRODUCT_FORM", "STREAM", "SUB_CLASS", "MEASURE"],
                                                           dropna=False, observed=True, as_index=False)["QTY_MT"].sum()
                                  if not resource_prod.empty else resource_prod),
        "inventory_weekly": aggregate_trend(trend, "WEEK", ["BU", "PLANT", "PRODUCT_FORM", "STREAM", "SUB_CLASS"]),
        "inventory_monthly": aggregate_trend(trend, "MONTH", ["BU", "PLANT", "PRODUCT_FORM", "STREAM", "SUB_CLASS"]),
        "fg_subclass_weekly": aggregate_trend(trend, "WEEK", ["BU", "PLANT", "PRODUCT_FORM", "STREAM", "SUB_CLASS"], ["Production", "Dispatch"]),
        "fg_subclass_monthly": aggregate_trend(trend, "MONTH", ["BU", "PLANT", "PRODUCT_FORM", "STREAM", "SUB_CLASS"], ["Production", "Dispatch"]),
        "material_flow_weekly": aggregate_trend(trend, "WEEK", ["BU", "PLANT", "RESOURCE", "PRODUCT_FORM", "STREAM", "SUB_CLASS"]),
        "material_flow_monthly": aggregate_trend(trend, "MONTH", ["BU", "PLANT", "RESOURCE", "PRODUCT_FORM", "STREAM", "SUB_CLASS"]),
        "intransit": aggregate_trend(trend, "WEEK", ["BU", "PLANT", "ITEM", "PRODUCT_FORM", "STREAM", "SUB_CLASS"], ["In-Transit"]),
        "bottlenecks": bottlenecks,
        "capacity": capacity if not capacity.empty else bottlenecks,
        "forecast_flow": flow,
        "forecast_flow_summary": forecast_summary,
        "bu_head_forecast_weekly": bu_head,
        "bu_head_forecast_monthly": (bu_head.groupby(["MONTH", "DFU", "DMDGROUP", "LOCATION", "BU", "MAPPED_ITEM", "SOURCE_MODE", "SOURCE_FILE"], dropna=False, observed=True, as_index=False)["BU_FCST_MT"].sum()
                                      if not bu_head.empty else pd.DataFrame(columns=["MONTH", "DFU", "DMDGROUP", "LOCATION", "BU", "MAPPED_ITEM", "SOURCE_MODE", "SOURCE_FILE", "BU_FCST_MT"])),
        "dfu_lost_mapping": loss,
        "dfu_lost_top5000": loss.nlargest(5000, "LOST_MT") if not loss.empty else loss,
        "weekly_balance_export": aggregate_trend(trend, "WEEK", ["BU", "PLANT", "PRODUCT_FORM"]),
        "monthly_balance_export": aggregate_trend(trend, "MONTH", ["BU", "PLANT", "PRODUCT_FORM"]),
        "service_summary": service_summary,
        "otif_summary": otif_summary,
        "mapping_column_matches": build_mapping_matches(sources),
        "forecast_reconciliation": build_forecast_reconciliation(sources, flow, loss),
        "diagnostics": build_diagnostics(sources),
        "forecast_waterfall": pd.DataFrame({"STAGE": ["BU_FCST_MT", "DFUTOSKU_FCST_MT", "NETTED_FCST_MT", "CUSTORDER_MT", "TOTAL_FCST_MT", "SCHED_QTY_MT"],
                                             "QTY_MT": [float(flow.get(c, pd.Series(dtype=float)).sum()) for c in ["BU_FCST_MT", "DFUTOSKU_FCST_MT", "NETTED_FCST_MT", "CUSTORDER_MT", "TOTAL_FCST_MT", "SCHED_QTY_MT"]]}),
        "service_demand": service,
        "value_at_risk": margin,
        "network_trace": network,
        "root_cause_misses": root_causes,
        "order_status": order_status,
        "inventory_health": inventory_health,
        "inventory_placeholders": inventory_placeholders,
        "overdue_purchase_orders": overdue_po,
        "sourcing": sourcing,
        "sourcing_issues": sourcing_issues,
        "master_health": master_health,
        "where_to_hit": where_to_hit,
        "item_master": item_master,
        "order_priority": order_priority,
        "plant_cards": build_plant_cards(trend, flow),
        "network_lanes_weekly": network_lanes,
        "inventory_balance_wide": inventory_balance_wide,
        "inventory_balance_weeks": pd.DataFrame({"WEEK_COLUMN": inventory_balance_weeks}),
        "schema_report": build_schema_report(sources),
        "simulation_catalog": pd.DataFrame([{"FILE": s.path.name, "SIMULATION": sim} for s in sources.values() for sim in s.available_simulations]),
    })
    if trend.empty: model.warnings.append("No weekly material-flow fact could be built. Check date and quantity column aliases in Schema Report.")
    if loss.empty: model.warnings.append("No DFU mapping loss was detected, or the BU forecast/mapping exports were not found.")
    return model


SHEET_MAP = [
    ("KPI Dashboard", None), ("Architecture", None),
    ("Where to Hit", "where_to_hit"), ("Resource Status", "capacity"),
    ("Forecast Reconciliation", "forecast_reconciliation"), ("Forecast Flow Summary", "forecast_flow_summary"),
    ("DFU Loss Top 5000", "dfu_lost_top5000"),
    ("Weekly Balance", "weekly_balance_export"), ("Monthly Balance", "monthly_balance_export"),
    ("Service Summary", "service_summary"), ("OTIF Summary", "otif_summary"),
    ("Plant Cards", "plant_cards"), ("Master Health", "master_health"),
    ("Mapping Matches", "mapping_column_matches"), ("Diagnostics", "diagnostics"),
    ("Schema Report", "schema_report"),
]


def _safe_excel_frame(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in out.columns:
        if pd.api.types.is_datetime64_any_dtype(out[col]):
            # XlsxWriter cannot write timezone-aware timestamps.
            try: out[col] = out[col].dt.tz_localize(None)
            except (TypeError, AttributeError): pass
    for col in out.select_dtypes(include=["number"]).columns:
        values = pd.to_numeric(out[col], errors="coerce")
        out[col] = values.mask(np.isinf(values), np.nan)
    return out


def _table_name(sheet: str) -> str:
    return "T_" + re.sub(r"[^A-Za-z0-9]", "", sheet)[:200]


def _xlsx_value(value: Any) -> Any:
    if value is None or value is pd.NaT:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    if isinstance(value, np.generic):
        return value.item()
    return value


def _write_frame_rows(ws: Any, df: pd.DataFrame, startrow: int, startcol: int, header_fmt: Any) -> None:
    """Write a frame in row order so XlsxWriter constant-memory mode is safe."""
    ws.write_row(startrow, startcol, [str(c) for c in df.columns], header_fmt)
    for row_number, values in enumerate(df.itertuples(index=False, name=None), startrow + 1):
        ws.write_row(row_number, startcol, [_xlsx_value(v) for v in values])


def excel_bytes(model: TwinModel) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="dd-mmm-yyyy", date_format="dd-mmm-yyyy",
                        engine_kwargs={"options": {"constant_memory": True, "nan_inf_to_errors": True}}) as writer:
        wb = writer.book
        navy, blue, pale, orange, red, green, white, grey = "#17365D", "#1F4E78", "#D9EAF7", "#F4B183", "#C00000", "#70AD47", "#FFFFFF", "#E7E6E6"
        title_fmt = wb.add_format({"bold": True, "font_size": 22, "font_color": white, "bg_color": navy, "align": "left", "valign": "vcenter"})
        subtitle_fmt = wb.add_format({"font_size": 10, "font_color": "#666666"})
        header_fmt = wb.add_format({"bold": True, "font_color": white, "bg_color": blue, "border": 1, "text_wrap": True})
        kpi_label = wb.add_format({"bold": True, "font_color": navy, "bg_color": pale, "border": 1})
        kpi_value = wb.add_format({"bold": True, "font_size": 16, "font_color": navy, "border": 1, "num_format": "#,##0.0"})
        nav_fmt = wb.add_format({"font_color": blue, "underline": True, "bg_color": pale, "border": 1})
        warn_fmt = wb.add_format({"font_color": red, "text_wrap": True})
        pct_fmt = wb.add_format({"num_format": "0.0%"})
        mt_fmt = wb.add_format({"num_format": "#,##0.0"})
        date_fmt = wb.add_format({"num_format": "dd-mmm-yyyy"})

        home = wb.add_worksheet("Home")
        writer.sheets["Home"] = home
        home.set_tab_color(navy); home.set_column("A:A", 4); home.set_column("B:B", 34); home.set_column("C:C", 70)
        home.set_row(1, 34); home.merge_range("B2:C2", "ESP Network Digital Twin", title_fmt)
        home.write("B3", f"Version {APP_VERSION} | Generated {model.generated_at:%d-%b-%Y %H:%M}", subtitle_fmt)
        home.write("B5", "Data folder", kpi_label); home.write("C5", str(model.folder))
        home.write("B6", "Simulation", kpi_label); home.write("C6", model.simulation or "All simulations")
        home.write("B8", "Workbook navigation", header_fmt); home.write("C8", "Purpose", header_fmt)
        purpose = {
            "KPI Dashboard": "Headline network, forecast and exception KPIs",
            "Architecture": "Data lineage, transformations and control logic",
            "Bottleneck Resources": "Resource load/capacity and >=95% utilization flags",
            "DFU Lost Mapping": "BU forecast lost through missing/failed DFU-to-SKU mapping",
            "Schema Report": "Loaded files, row counts, simulation filtering and columns",
        }
        row = 8
        for sheet, _ in SHEET_MAP:
            home.write_url(row, 1, f"internal:'{sheet}'!A1", nav_fmt, sheet)
            home.write(row, 2, purpose.get(sheet, sheet.replace("Fact ", "Detailed fact: ")))
            row += 1
        if model.warnings:
            row += 1; home.write(row, 1, "Load warnings", header_fmt); row += 1
            for warning in model.warnings[:100]: home.merge_range(row, 1, row, 2, warning, warn_fmt); row += 1
        home.freeze_panes(8, 1)

        # Dashboard
        dash = wb.add_worksheet("KPI Dashboard"); writer.sheets["KPI Dashboard"] = dash
        dash.set_column("A:A", 3); dash.set_column("B:H", 20); dash.set_row(1, 34)
        dash.merge_range("B2:H2", "Network KPI Dashboard", title_fmt)
        dash.write_url("B3", "internal:'Home'!B2", nav_fmt, "← Home")
        kpi_items = list(model.kpis().items())
        for group_start in range(0, len(kpi_items), 4):
            r = 4 + (group_start // 4) * 4
            group = kpi_items[group_start:group_start + 4]
            for position, (label, _) in enumerate(group):
                c = 1 + position * 2; dash.write(r, c, label, kpi_label); dash.set_column(c, c, 20)
            for position, (_, value) in enumerate(group):
                c = 1 + position * 2
                safe_value = "n/a" if isinstance(value, (float, np.floating)) and not np.isfinite(value) else value
                dash.write(r + 1, c, safe_value, kpi_value)

        # Put a compact monthly trend below dashboard and chart it.
        monthly = aggregate_trend(model.tables["weekly_trend"], "MONTH", [], ["Production", "Dispatch", "Closing Inventory"])
        if not monthly.empty:
            monthly_p = monthly.pivot_table(index="MONTH", columns="MEASURE", values="QTY_MT", aggfunc="sum", fill_value=0).reset_index()
            start = 20
            dash.set_column(1, 1, 14, date_fmt)
            _write_frame_rows(dash, monthly_p, startrow=start, startcol=1, header_fmt=header_fmt)
            chart = wb.add_chart({"type": "line"})
            for c in range(1, len(monthly_p.columns)):
                chart.add_series({"name": ["KPI Dashboard", start, 1+c], "categories": ["KPI Dashboard", start+1, 1, start+len(monthly_p), 1],
                                  "values": ["KPI Dashboard", start+1, 1+c, start+len(monthly_p), 1+c], "line": {"width": 2.25}})
            chart.set_title({"name": "Monthly Network Trend (MT)"}); chart.set_y_axis({"name": "MT"}); chart.set_legend({"position": "bottom"})
            dash.insert_chart("I4", chart, {"x_scale": 1.25, "y_scale": 1.15})

        arch = wb.add_worksheet("Architecture"); writer.sheets["Architecture"] = arch
        arch.set_column("A:A", 4); arch.set_column("B:B", 25); arch.set_column("C:C", 45); arch.set_column("D:D", 70)
        arch.merge_range("B2:D2", "ESP Digital Twin Architecture & Controls", title_fmt)
        architecture = [
            ("1. Ingest", "CSV/TXT exports in one folder", "Recursive load; encoding and delimiter detection; columns normalised"),
            ("2. Scenario", "SIMULATION_NAME fields only", "Exact simulation filter. OPTIONSET and generic scenario/plan/run fields are excluded."),
            ("3. Quantity", "ESP quantity fields", "kg converted to MT (/1000); explicit *_MT fields are retained"),
            ("4. Product", "U_PRODUCTLVL_1, U_SUB_CLASS", "Sub-class = Container when U_PRODUCTLVL_1='E'; otherwise U_SUB_CLASS"),
            ("5. Time", "Transaction/bucket date", "Monday week start and calendar month start"),
            ("6. Forecast", "BU_FCST → DFUTOSKU → NETTED + Orders → Total → Schedule", "Mapping loss and MET % calculated"),
            ("7. Deliver", "Streamlit + Excel", "Same model feeds interactive views, workbook and CLI"),
        ]
        arch.write_row(4, 1, ["Layer", "Input", "Rule / Output"], header_fmt)
        for r, values in enumerate(architecture, 5): arch.write_row(r, 1, values)
        arch.write_url("B14", "internal:'Home'!B2", nav_fmt, "← Home")

        # Data sheets
        sheet_aliases = {"Fact Forecast Flow": "Demand vs Supply", "Fact DFU Loss": "DFU Lost Mapping"}
        for sheet, key in SHEET_MAP:
            if sheet in {"KPI Dashboard", "Architecture"}: continue
            df = _safe_excel_frame(model.tables.get(key or "", pd.DataFrame()))
            if len(df) > EXCEL_MAX_DATA_ROWS:
                model.warnings.append(f"{sheet}: truncated from {len(df):,} to {EXCEL_MAX_DATA_ROWS:,} rows for Excel.")
                df = df.iloc[:EXCEL_MAX_DATA_ROWS].copy()
            ws = wb.add_worksheet(sheet); writer.sheets[sheet] = ws
            ws.set_tab_color(blue if not sheet.startswith("Fact ") else grey)
            ws.set_row(0, 30); ws.merge_range(0, 0, 0, max(2, len(df.columns)-1), sheet, title_fmt)
            ws.write_url(1, 0, "internal:'Home'!B2", nav_fmt, "← Home")
            if sheet in sheet_aliases:
                target = sheet_aliases[sheet]
                ws.write_url(3, 0, f"internal:'{target}'!A1", nav_fmt, f"Detailed fact is stored once in {target} to keep the workbook memory-efficient.")
                continue
            if len(df.columns):
                for c, col in enumerate(df.columns):
                    display_values = df[col].head(200).astype(object).map(str)
                    width = min(42, max([11, len(str(col)) + 2, *display_values.map(len).tolist()]))
                    fmt = mt_fmt if str(col).endswith("_MT") or str(col) == "QTY_MT" else None
                    if pd.api.types.is_datetime64_any_dtype(df[col]): fmt = date_fmt
                    if str(col).endswith("_PCT"): fmt = wb.add_format({"num_format": "0.0"})
                    ws.set_column(c, c, width, fmt)
            _write_frame_rows(ws, df, startrow=3, startcol=0, header_fmt=header_fmt)
            ws.freeze_panes(4, 1); ws.autofilter(3, 0, max(3, len(df)+3), max(0, len(df.columns)-1))
            if len(df.columns):
                # Highlight exceptions.
                if "LOST_MT" in df.columns and len(df):
                    c = df.columns.get_loc("LOST_MT"); ws.conditional_format(4, c, len(df)+3, c, {"type": "data_bar", "bar_color": orange})
                if "BOTTLENECK" in df.columns and len(df):
                    c = df.columns.get_loc("BOTTLENECK"); ws.conditional_format(4, c, len(df)+3, c, {"type": "text", "criteria": "containing", "value": "TRUE", "format": wb.add_format({"bg_color": "#FFC7CE", "font_color": red})})
                if "MET_PCT" in df.columns and len(df):
                    c = df.columns.get_loc("MET_PCT"); ws.conditional_format(4, c, len(df)+3, c, {"type": "3_color_scale", "min_color": "#F8696B", "mid_color": "#FFEB84", "max_color": "#63BE7B"})
        wb.set_properties({"title": "ESP Network Digital Twin", "subject": "Blue Yonder ESP network analysis", "author": "ESP Digital Twin", "comments": f"Generated by twin_engine.py {APP_VERSION}"})
    return output.getvalue()


def _excel_sheet_name(name: str) -> str:
    return re.sub(r"[\[\]:*?/\\]", "_", name)[:31]


def _write_inventory_drilldown(ws: Any, wb: Any, wide: pd.DataFrame, week_labels: list[str], home_fmt: Any) -> None:
    """Write the formula-driven 1/2/3-level inventory balance sheet."""
    import inv_balance_report as ibr
    from xlsxwriter.utility import xl_col_to_name
    df = ibr.ordered(wide, week_labels)
    headers = [c for c in df.columns if c != "_lvl"]
    navy, orange, white = "#17365D", "#F47732", "#FFFFFF"
    title = wb.add_format({"font_name": "Arial", "bold": True, "font_size": 15, "font_color": navy})
    note = wb.add_format({"font_name": "Arial", "italic": True, "font_color": "#5B6B7A"})
    head = wb.add_format({"font_name": "Arial", "bold": True, "font_color": white, "bg_color": navy, "border": 1, "align": "center", "text_wrap": True})
    horizon_head = wb.add_format({"font_name": "Arial", "bold": True, "font_color": white, "bg_color": orange, "border": 1, "align": "center"})
    num = wb.add_format({"font_name": "Arial", "num_format": '#,##0.0;[Red]-#,##0.0;"-"'})
    text = wb.add_format({"font_name": "Arial"})
    close_num = wb.add_format({"font_name": "Arial", "bold": True, "bg_color": "#E6EEF1", "num_format": '#,##0.0;[Red]-#,##0.0;"-"'})
    tints = {"RM": "#FDEBD0", "WIP": "#FCF3CF", "FG": "#D5F5E3", "Scrap": "#FADBD8", "Other": "#E8EEF3", "All": "#D9EAF7"}
    sum_text = {k: wb.add_format({"font_name": "Arial", "bold": True, "bg_color": v}) for k, v in tints.items()}
    sum_num = {k: wb.add_format({"font_name": "Arial", "bold": True, "bg_color": v, "num_format": '#,##0.0;[Red]-#,##0.0;"-"'}) for k, v in tints.items()}
    ws.write(0, 0, "Inventory balance — weekly drill-down (MT): Plant → Category → Resource / Form / Sub-class or Container", title)
    ws.write(1, 0, "Use outline levels 1/2/3 or [+]/[-]. Closing = Opening + Production + Receipts − Consumption − both Dispatch rows. Next-week Opening = prior Closing. Red = negative.", note)
    ws.write_url(0, max(7, len(headers)-1), "internal:'Home'!B2", home_fmt, "⌂ Home")
    for c, h in enumerate(headers): ws.write(3, c, h, horizon_head if h == "Horizon" else head)
    widths = [26, 9, 10, 11, 24, 37] + [11] * (len(week_labels) + 1)
    for c, width in enumerate(widths): ws.set_column(c, c, width)
    block_rows: dict[tuple, dict[str, int]] = {}
    id_cols = ["Resource", "Plant", "Category", "Form", "Sub-class / Container"]
    for i, rec in df.iterrows():
        block_rows.setdefault(tuple(rec[c] for c in id_cols), {})[rec["Measure"]] = 5 + i
    for i, rec in df.iterrows():
        row = 4 + i; level = int(rec["_lvl"]); category = str(rec["Category"])
        color_key = category if category in tints else "Other"
        ws.set_row(row, None, None, {"level": level, "hidden": level > 1})
        summary = level <= 2
        for c in range(6): ws.write(row, c, _xlsx_value(rec[headers[c]]), sum_text[color_key] if summary else text)
        rows_for = block_rows[tuple(rec[c] for c in id_cols)]
        measure = rec["Measure"]
        for c in range(6, len(headers)):
            fmt = sum_num[color_key] if summary else (close_num if measure == "Closing inventory" else num)
            if c < 6 + len(week_labels):
                wi = c - 6; colx = xl_col_to_name(c)
                if measure == "Opening inventory" and wi > 0:
                    formula = f"={xl_col_to_name(c-1)}{rows_for['Closing inventory']}"
                    ws.write_formula(row, c, formula, fmt)
                elif measure == "Closing inventory":
                    formula = (f"={colx}{rows_for['Opening inventory']}+{colx}{rows_for['Production']}+"
                               f"{colx}{rows_for['Receipts (transfer/purchase/sched.)']}-"
                               f"{colx}{rows_for['Consumption (to next stage)']}-"
                               f"{colx}{rows_for['Dispatch - customer/zone']}-"
                               f"{colx}{rows_for['Dispatch - inter-plant transfer']}")
                    ws.write_formula(row, c, formula, fmt)
                else:
                    ws.write_number(row, c, float(rec[headers[c]]), fmt)
            else:
                first_col, last_col = xl_col_to_name(6), xl_col_to_name(5 + len(week_labels))
                if measure == "Opening inventory": formula = f"={first_col}{row+1}"
                elif measure in ("Closing inventory", "In-transit to plant (closing)"): formula = f"={last_col}{row+1}"
                else: formula = f"=SUM({first_col}{row+1}:{last_col}{row+1})"
                ws.write_formula(row, c, formula, fmt)
    ws.freeze_panes(4, 6); ws.autofilter(3, 0, max(3, len(df)+3), max(0, len(headers)-1))
    ws.outline_settings(True, False, True, False)
    if len(df):
        ws.conditional_format(4, 6, len(df)+3, len(headers)-1, {"type": "cell", "criteria": "<", "value": 0,
                                                                     "format": wb.add_format({"font_name": "Arial", "font_color": "#C00000"})})


def excel_bytes(model: TwinModel) -> bytes:
    """Create the shareable v2.5 workbook in the supplied reference style."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="dd-mmm-yyyy", date_format="dd-mmm-yyyy",
                        engine_kwargs={"options": {"constant_memory": True, "nan_inf_to_errors": True}}) as writer:
        wb = writer.book; wb.set_calc_mode("auto")
        navy, blue, pale, orange, red, green, white, grey = "#17365D", "#1F4E78", "#D9EAF7", "#F47732", "#C00000", "#70AD47", "#FFFFFF", "#E7E6E6"
        base = wb.add_format({"font_name": "Arial"})
        title = wb.add_format({"font_name": "Arial", "bold": True, "font_size": 20, "font_color": white, "bg_color": navy, "valign": "vcenter"})
        subtitle = wb.add_format({"font_name": "Arial", "font_size": 10, "font_color": "#5B6B7A"})
        header = wb.add_format({"font_name": "Arial", "bold": True, "font_color": white, "bg_color": blue, "border": 1, "text_wrap": True})
        home_link = wb.add_format({"font_name": "Arial", "font_color": blue, "underline": True, "bg_color": pale, "border": 1})
        label = wb.add_format({"font_name": "Arial", "bold": True, "font_color": navy, "bg_color": pale, "border": 1})
        value = wb.add_format({"font_name": "Arial", "bold": True, "font_size": 15, "font_color": navy, "border": 1, "num_format": '#,##0.0'})
        mt_fmt = wb.add_format({"font_name": "Arial", "num_format": '#,##0.0;[Red]-#,##0.0;"-"'})
        date_fmt = wb.add_format({"font_name": "Arial", "num_format": "dd-mmm-yyyy"})

        sheets: list[tuple[str, str | None]] = [
            ("Architecture", None), ("KPI_Dashboard", None), ("Network_Map", None),
            ("Bottleneck_Resources", "capacity"), ("Resource_Prod_Weekly", "resource_prod_weekly"),
            ("Resource_Prod_Monthly", "resource_prod_monthly"), ("Inventory_Balance_Weekly", None),
            ("Inventory_Balance_Monthly", "inventory_monthly"), ("FG_SubClass_Weekly", "fg_subclass_weekly"),
            ("FG_SubClass_Monthly", "fg_subclass_monthly"), ("Material_Flow_Weekly", "material_flow_weekly"),
            ("Material_Flow_Monthly", "material_flow_monthly"), ("Demand_vs_Supply", "forecast_flow_summary"),
            ("In_Transit", "intransit"), ("Order_Priority", "order_priority"), ("DFU_Loss", "dfu_lost_mapping"),
            ("OTIF", "order_status"), ("ITEM_MASTER", "item_master"), ("Assumptions", "forecast_reconciliation"),
            ("Fact_Wide", "inventory_balance_wide"), ("Fact_Long", "weekly_trend"),
            ("ITEM_MASTER_Dynamic", "item_master"), ("Data_Quality", "master_health"),
        ]
        plant_names = sorted(x for x in model.tables.get("plant_cards", pd.DataFrame()).get("PLANT", pd.Series(dtype=str)).dropna().astype(str).unique() if x and x != "(Blank)")

        # Home / navigation.
        home = wb.add_worksheet("Home"); writer.sheets["Home"] = home
        home.set_tab_color(navy); home.set_column("A:A", 3); home.set_column("B:B", 34); home.set_column("C:C", 80)
        home.set_row(1, 34); home.merge_range("B2:C2", "ESP Network Digital Twin — Blue Yonder ESP", title)
        home.write("B3", f"Version {APP_VERSION} | Generated {model.generated_at:%d-%b-%Y %H:%M} | kg → MT", subtitle)
        home.write("B5", "Data folder", label); home.write("C5", str(model.folder), base)
        home.write("B6", "Simulation", label); home.write("C6", model.simulation or "All simulations", base)
        home.write_row(7, 1, ["Workbook navigation", "Purpose"], header)
        purposes = {
            "Architecture": "Animated weekly network architecture and control lineage", "KPI_Dashboard": "Headline network KPIs",
            "Network_Map": "Plant/demand-zone map and ESP_PLANARRIV lanes", "Inventory_Balance_Weekly": "1/2/3-level formula drill-down",
            "Bottleneck_Resources": "Resource status and machine-centre detail", "Demand_vs_Supply": "Forecast/supply reconciliation",
        }
        r = 8
        for sheet, _ in sheets:
            home.write_url(r, 1, f"internal:'{sheet}'!A1", home_link, sheet); home.write(r, 2, purposes.get(sheet, sheet.replace("_", " ")), base); r += 1
        for plant in plant_names:
            name = _excel_sheet_name(f"Plant_{plant}")
            home.write_url(r, 1, f"internal:'{name}'!A1", home_link, name); home.write(r, 2, f"Weekly production, inventory and flow for {plant}", base); r += 1
        home.freeze_panes(8, 1)

        gif_bytes = network_animation_gif(model)
        from PIL import Image
        first_png = io.BytesIO(); gif_image = Image.open(io.BytesIO(gif_bytes)); gif_image.seek(0); gif_image.save(first_png, format="PNG"); first_png.seek(0)

        # Architecture with embedded animated GIF.
        arch = wb.add_worksheet("Architecture"); writer.sheets["Architecture"] = arch
        arch.set_column("A:A", 3); arch.set_column("B:B", 24); arch.set_column("C:C", 42); arch.set_column("D:D", 70)
        arch.merge_range("B2:D2", "ESP Digital Twin Architecture & Weekly Network Clip", title); arch.write_url("B3", "internal:'Home'!B2", home_link, "← Home")
        architecture = [
            ("1. Ingest", "ESP CSV exports", "Latest exact source files; tolerant schema aliases"),
            ("2. Horizon", "ESP_INVENTORY.AVAILDATE", "First weekly bucket; historical dates are excluded"),
            ("3. Quantity", "ESP quantity fields", "kg / 1,000 → MT; explicit *_MT retained"),
            ("4. Product", "U_PRODUCTLVL_1/U_SUB_CLASS", "E uses Container n; otherwise U_SUB_CLASS"),
            ("5. Inventory", "Opening + flows", "Formula Closing and carry-forward Opening; negatives highlighted"),
            ("6. OTIF", "ESP_CUSTORDERSTATUS CUST", "Pre-plan SHIPDATE excluded; met/completed and SCHEDDATE ≤ SHIPDATE; Qty-MT weighted"),
            ("7. DFU loss", "UDT → FE → ESP", "Demand DMDUNIT maps to supply ITEM; supply joins ITEM + DFULOC + month only"),
            ("8. Deliver", "Streamlit + Excel", "One model feeds app, clip and workbook"),
        ]
        arch.write_row(4, 1, ["Layer", "Input", "Rule / Output"], header)
        for rr, rowvals in enumerate(architecture, 5): arch.write_row(rr, 1, rowvals, base)
        arch.write(14, 1, "Animated weekly material flow", header)
        arch.insert_image("B16", "ESP_weekly_network.gif", {"image_data": io.BytesIO(gif_bytes), "x_scale": 0.70, "y_scale": 0.70})

        # KPI dashboard.
        dash = wb.add_worksheet("KPI_Dashboard"); writer.sheets["KPI_Dashboard"] = dash
        dash.set_column("A:A", 3); dash.set_column("B:I", 21); dash.set_row(1, 34)
        dash.merge_range("B2:I2", "Network KPI Dashboard", title); dash.write_url("B3", "internal:'Home'!B2", home_link, "← Home")
        items = list(model.kpis().items())
        for i, (k, v) in enumerate(items):
            rr = 4 + (i // 4) * 3; cc = 1 + (i % 4) * 2
            dash.write(rr, cc, k, label); dash.write(rr+1, cc, "n/a" if isinstance(v, float) and not np.isfinite(v) else v, value)
        trend = model.tables.get("weekly_trend", pd.DataFrame())
        weekly = (trend.groupby(["WEEK", "MEASURE"], observed=True, as_index=False)["QTY_MT"].sum().pivot(index="WEEK", columns="MEASURE", values="QTY_MT").fillna(0).reset_index()
                  if not trend.empty else pd.DataFrame())
        if not weekly.empty:
            startrow = 18; dash.write_row(startrow, 1, list(weekly.columns), header)
            for i, rowvals in enumerate(weekly.itertuples(index=False, name=None), startrow+1): dash.write_row(i, 1, [_xlsx_value(x) for x in rowvals], base)
            chart = wb.add_chart({"type": "line"})
            for c in range(1, len(weekly.columns)):
                chart.add_series({"name": ["KPI_Dashboard", startrow, 1+c], "categories": ["KPI_Dashboard", startrow+1, 1, startrow+len(weekly), 1],
                                  "values": ["KPI_Dashboard", startrow+1, 1+c, startrow+len(weekly), 1+c]})
            chart.set_title({"name": "Weekly material-flow trend from ESP_INVENTORY AVAILDATE"}); chart.set_legend({"position": "bottom"}); dash.insert_chart("J4", chart)

        # Network map (static first frame + lane table).
        net = wb.add_worksheet("Network_Map"); writer.sheets["Network_Map"] = net
        net.merge_range("B2:L2", "Hindalco Supply Network Digital Twin — Network Map", title); net.write_url("B3", "internal:'Home'!B2", home_link, "← Home")
        net.insert_image("B5", "ESP_weekly_network.png", {"image_data": first_png, "x_scale": .72, "y_scale": .72})
        lanes = _safe_excel_frame(model.tables.get("network_lanes_weekly", pd.DataFrame()))
        net.write_row(34, 1, list(lanes.columns), header)
        for i, rowvals in enumerate(lanes.head(50000).itertuples(index=False, name=None), 35): net.write_row(i, 1, [_xlsx_value(x) for x in rowvals], base)
        net.freeze_panes(35, 1)

        # Inventory formula hierarchy.
        inv = model.tables.get("inventory_balance_wide", pd.DataFrame())
        week_labels = model.tables.get("inventory_balance_weeks", pd.DataFrame()).get("WEEK_COLUMN", pd.Series(dtype=str)).astype(str).tolist()
        inv_ws = wb.add_worksheet("Inventory_Balance_Weekly"); writer.sheets["Inventory_Balance_Weekly"] = inv_ws
        if not inv.empty and week_labels: _write_inventory_drilldown(inv_ws, wb, inv, week_labels, home_link)
        else: inv_ws.write("A1", "Inventory balance could not be built; see Data_Quality.", base)

        # Generic styled data sheets, preserving the reference names.
        skip = {"Architecture", "KPI_Dashboard", "Network_Map", "Inventory_Balance_Weekly"}
        for sheet, key in sheets:
            if sheet in skip: continue
            df = _safe_excel_frame(model.tables.get(key or "", pd.DataFrame()))
            if len(df) > EXCEL_MAX_DATA_ROWS: df = df.iloc[:EXCEL_MAX_DATA_ROWS]
            ws = wb.add_worksheet(sheet); writer.sheets[sheet] = ws
            ws.set_tab_color(grey if sheet.startswith("Fact_") else blue)
            end_col = max(2, len(df.columns)-1); ws.merge_range(0, 0, 0, end_col, sheet.replace("_", " "), title)
            ws.write_url(1, 0, "internal:'Home'!B2", home_link, "← Home")
            ws.write_row(3, 0, list(df.columns), header)
            for rr, rowvals in enumerate(df.itertuples(index=False, name=None), 4): ws.write_row(rr, 0, [_xlsx_value(x) for x in rowvals], base)
            for c, col in enumerate(df.columns):
                sample = df[col].head(150).astype(object).map(str)
                width = min(42, max([11, len(str(col))+2, *sample.map(len).tolist()]))
                fmt = date_fmt if pd.api.types.is_datetime64_any_dtype(df[col]) else (mt_fmt if str(col).endswith("_MT") or str(col) in {"QTY_MT", "Horizon"} else base)
                ws.set_column(c, c, width, fmt)
            ws.freeze_panes(4, 1); ws.autofilter(3, 0, max(3, len(df)+3), max(0, len(df.columns)-1))
            if "LOST_MT" in df.columns and len(df):
                c = df.columns.get_loc("LOST_MT"); ws.conditional_format(4, c, len(df)+3, c, {"type": "data_bar", "bar_color": orange})
            if "BOTTLENECK" in df.columns and len(df):
                c = df.columns.get_loc("BOTTLENECK"); ws.conditional_format(4, c, len(df)+3, c, {"type": "text", "criteria": "containing", "value": "TRUE", "format": wb.add_format({"font_name": "Arial", "bg_color": "#FFC7CE", "font_color": red})})

        # Plant-level drill-down sheets.
        flow = model.tables.get("material_flow_weekly", pd.DataFrame())
        for plant in plant_names:
            name = _excel_sheet_name(f"Plant_{plant}"); ws = wb.add_worksheet(name); writer.sheets[name] = ws
            df = _safe_excel_frame(flow[flow.get("PLANT", pd.Series(index=flow.index, dtype=str)).astype(str).eq(plant)].copy()) if not flow.empty else pd.DataFrame()
            ws.merge_range(0, 0, 0, max(2, len(df.columns)-1), f"Plant {plant} — weekly resource and material flow", title); ws.write_url(1, 0, "internal:'Home'!B2", home_link, "← Home")
            ws.write_row(3, 0, list(df.columns), header)
            for rr, rowvals in enumerate(df.itertuples(index=False, name=None), 4): ws.write_row(rr, 0, [_xlsx_value(x) for x in rowvals], base)
            ws.freeze_panes(4, 1); ws.autofilter(3, 0, max(3, len(df)+3), max(0, len(df.columns)-1))
        wb.set_properties({"title": "ESP Network Digital Twin", "subject": "Blue Yonder ESP network analysis", "author": "ESP Digital Twin", "comments": f"Generated by twin_engine.py {APP_VERSION}"})
    return output.getvalue()


def write_excel(model: TwinModel, output_path: str | Path) -> Path:
    path = Path(output_path).expanduser()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(excel_bytes(model))
    return path


def write_findings(model: TwinModel, output_path: str | Path) -> Path:
    """Write a compact, timestamp-safe findings workbook."""
    path = Path(output_path).expanduser(); path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="xlsxwriter", datetime_format="dd-mmm-yyyy") as writer:
        pd.DataFrame({"KPI": list(model.kpis()), "Value": list(model.kpis().values())}).to_excel(writer, sheet_name="KPIs", index=False)
        model.tables.get("forecast_reconciliation", pd.DataFrame()).to_excel(writer, sheet_name="Forecast_Reconciliation", index=False)
        model.tables.get("diagnostics", pd.DataFrame()).to_excel(writer, sheet_name="Diagnostics", index=False)
        model.tables.get("master_health", pd.DataFrame()).to_excel(writer, sheet_name="Data_Quality", index=False)
        for ws in writer.sheets.values():
            ws.set_column(0, 0, 34); ws.set_column(1, 20, 22)
    return path


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Build the ESP Network Digital Twin Excel workbook.")
    parser.add_argument("--folder", required=True, help="Folder containing ESP CSV/TXT exports")
    parser.add_argument("--simulation", default=None, help="Simulation/scenario to retain, e.g. MONTHLY_WHATIF")
    parser.add_argument("--output", default=None, help="Optional output base name; files are always saved in APP_OUTPUT_FILES")
    args = parser.parse_args(argv)
    model = build_twin(args.folder, args.simulation)
    stem = Path(args.output).stem if args.output else f"ESP_Network_Digital_Twin_{args.simulation or 'ALL'}"
    path = timestamped_output_path(args.folder, stem, ".xlsx")
    write_excel(model, path)
    gif_path = timestamped_output_path(args.folder, "ESP_Weekly_Network_Animation", ".gif")
    gif_path.write_bytes(network_animation_gif(model))
    findings_path = timestamped_output_path(args.folder, f"Digital_Twin_Findings_{args.simulation or 'ALL'}", ".xlsx")
    write_findings(model, findings_path)
    print(f"Created: {path.resolve()}")
    print(f"Created: {gif_path.resolve()}")
    print(f"Created: {findings_path.resolve()}")
    print(f"Loaded {len(model.sources)} files / {model.source_rows:,} rows")
    for warning in model.warnings: print(f"WARNING: {warning}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
