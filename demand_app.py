"""Streamlit Demand Management checks application.

Run:
    streamlit run demand_app.py
"""
from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st

from demand_engine import DM_VERSION, DemandModel, build_demand_model, demand_excel_bytes, write_demand_excel

DEFAULT_FOLDER = r"C:\Users\rakesh.ranjan\OneDrive - Aditya Birla Group\python projects\by ibp tests v2.5\ESP_RUN_FILES"

st.set_page_config(page_title="Demand Management Checks", page_icon="📈", layout="wide")


@st.cache_data(show_spinner="Loading demand-planning exports…", max_entries=4)
def load_model(folder: str, simulation: str, signature: tuple) -> DemandModel:
    del signature
    return build_demand_model(folder, simulation.strip() or None)


@st.cache_data(show_spinner="Building Demand Management workbook…", max_entries=4)
def workbook_bytes(folder: str, simulation: str, signature: tuple) -> bytes:
    return demand_excel_bytes(load_model(folder, simulation, signature))


def signature_for(folder: str) -> tuple:
    root = Path(folder).expanduser()
    if not root.exists(): return ()
    return tuple(sorted((str(p.relative_to(root)), p.stat().st_size, p.stat().st_mtime_ns)
                        for p in root.rglob("*") if p.is_file() and p.suffix.lower() in {".csv", ".txt"}))


def options(df: pd.DataFrame, column: str) -> list[str]:
    return sorted(df[column].dropna().astype(str).unique()) if column in df else []


def filter_frame(df: pd.DataFrame, selected: dict[str, list[str]]) -> pd.DataFrame:
    out = df
    for column, values in selected.items():
        if values and column in out: out = out[out[column].astype(str).isin(values)]
    return out


def table(df: pd.DataFrame, key: str, height: int = 460) -> None:
    st.dataframe(df, width="stretch", hide_index=True, height=height, key=key)


def mt(value: float) -> str:
    return f"{value:,.1f} MT"


st.title("📈 Demand Management Checks")
st.caption(f"Demand master, transaction and output controls · v{DM_VERSION}")

with st.sidebar:
    st.header("Data setup")
    folder = st.text_input("Folder containing CSV/TXT files", value=DEFAULT_FOLDER)
    simulation = st.text_input("Simulation (optional)", value="",
                               help="Filters SIMULATION/SCENARIO/PLAN/RUN only. OPTIONSET remains a measure.")
    if st.button("Reload changed files", width="stretch"):
        st.cache_data.clear()

signature = signature_for(folder)
if not signature:
    st.info("Enter a folder that contains demand-planning CSV/TXT exports.")
    st.stop()
try:
    model = load_model(folder, simulation, signature)
except Exception as exc:
    st.error(f"Could not run Demand Management checks: {exc}")
    st.stop()

flow = model.tables["forecast_flow"]
with st.sidebar:
    st.divider(); st.header("Analysis filters")
    selected = {}
    for column, label in [("BU", "BU"), ("PLANT", "Location / plant"), ("ITEM", "Item / SKU"), ("DFU", "DFU")]:
        vals = options(flow, column)
        selected[column] = st.multiselect(label, vals, placeholder="All") if vals else []

k = model.kpis()
cols = st.columns(7)
cols[0].metric("Files", f"{k['Files']:,}")
cols[1].metric("Rows", f"{k['Rows']:,}")
cols[2].metric("Quality issues", f"{k['Quality issues']:,}")
cols[3].metric("High issues", f"{k['High issues']:,}")
cols[4].metric("Mapping lost", mt(float(k["Mapping lost MT"])))
cols[5].metric("Weighted MET", f"{k['Weighted MET %']:.1f}%" if not np.isnan(k["Weighted MET %"]) else "n/a")
cols[6].metric("Actions", f"{k['Where-to-hit actions']:,}")

if model.warnings:
    with st.expander(f"Coverage notes ({len(model.warnings)})"):
        for warning in model.warnings: st.warning(warning)

(tab_hit, tab_files, tab_quality, tab_flow, tab_compare, tab_mapping, tab_export) = st.tabs([
    "Where to hit", "File coverage", "Data quality", "Forecast flow & netting",
    "Forecast vs orders / bias", "DFU mapping", "Excel export",
])

with tab_hit:
    st.subheader("Prioritised action list")
    hit = model.tables["where_to_hit"]
    priorities = sorted(hit["PRIORITY"].unique().tolist()) if not hit.empty else []
    areas = options(hit, "AREA")
    c1, c2 = st.columns(2)
    selected_priorities = c1.multiselect("Priority", priorities, default=priorities)
    selected_areas = c2.multiselect("Area", areas, default=areas)
    if selected_priorities: hit = hit[hit["PRIORITY"].isin(selected_priorities)]
    if selected_areas: hit = hit[hit["AREA"].isin(selected_areas)]
    table(hit, "where_to_hit", 590)

with tab_files:
    st.subheader("Master, transaction and output coverage")
    inventory = model.tables["file_inventory"]
    role_summary = (inventory.groupby("ROLE", observed=True, as_index=False)
                    .agg(FILES=("FILE", "count"), ROWS=("ROWS", "sum"), MIN_DATE=("MIN_DATE", "min"), MAX_DATE=("MAX_DATE", "max")))
    table(role_summary, "role_summary", 280)
    table(inventory, "file_inventory", 440)

with tab_quality:
    st.subheader("Data-quality checks")
    quality = model.tables["data_quality"]
    c1, c2, c3 = st.columns(3)
    sev = options(quality, "SEVERITY")
    roles = options(quality, "ROLE")
    cats = options(quality, "CATEGORY")
    sel_sev = c1.multiselect("Severity", sev, default=sev)
    sel_roles = c2.multiselect("Role", roles, default=roles)
    sel_cats = c3.multiselect("Check", cats, default=cats)
    if sel_sev: quality = quality[quality["SEVERITY"].astype(str).isin(sel_sev)]
    if sel_roles: quality = quality[quality["ROLE"].isin(sel_roles)]
    if sel_cats: quality = quality[quality["CATEGORY"].isin(sel_cats)]
    table(quality, "quality", 560)

with tab_flow:
    st.subheader("Forecast flow and netting")
    st.caption("BU_FCST → DFUTOSKU_FCST → NETTED_FCST + CUSTORDER → TOTAL_FCST → SCHED_QTY → MET %")
    f = filter_frame(flow, selected)
    if f.empty:
        st.info("No forecast-flow rows match the current filters or no forecast outputs were identified.")
    else:
        metric_cols = st.columns(6)
        metrics = [("BU forecast", "BU_FCST_MT"), ("SKU forecast", "DFUTOSKU_FCST_MT"),
                   ("Netted forecast", "NETTED_FCST_MT"), ("Customer orders", "CUSTORDER_MT"),
                   ("Total forecast", "TOTAL_FCST_MT"), ("Scheduled", "SCHED_QTY_MT")]
        for col, (label, field) in zip(metric_cols, metrics): col.metric(label, mt(float(f[field].sum())))
        monthly = f.groupby("MONTH", observed=True)[[field for _, field in metrics]].sum().sort_index()
        st.line_chart(monthly, height=390, width="stretch", y_label="MT", x_label="Month")
        table(f, "flow", 470)

with tab_compare:
    st.subheader("Forecast vs customer orders, fulfilment and actuals")
    fvo = filter_frame(model.tables["forecast_vs_orders"], selected)
    table(fvo, "forecast_vs_orders", 480)
    st.subheader("Bias and coverage where actuals allow")
    st.caption("Bias = forecast − actual. WAPE = absolute error / absolute actual. Bucket coverage compares forecasted and actual-active month buckets.")
    bias = filter_frame(model.tables["bias_coverage"], selected)
    table(bias, "bias_coverage", 390)

with tab_mapping:
    st.subheader("DFUs that lost forecast through DFU-to-SKU mapping")
    loss = filter_frame(model.tables["dfu_mapping_loss"], selected)
    if not loss.empty:
        c1, c2 = st.columns(2)
        c1.metric("Affected DFU-month rows", f"{len(loss):,}")
        c2.metric("Lost forecast", mt(float(loss["LOST_MT"].sum())))
        issue_opts = options(loss, "ISSUE")
        issue = st.multiselect("Issue type", issue_opts, default=issue_opts)
        if issue: loss = loss[loss["ISSUE"].isin(issue)]
    table(loss, "mapping_loss", 570)

with tab_export:
    st.subheader("Excel output")
    st.write("The workbook contains a navigable executive summary, file inventory, data-quality checks, forecast flow, forecast-vs-orders, bias/coverage, DFU mapping loss and the where-to-hit list.")
    data = workbook_bytes(folder, simulation, signature)
    file_name = "Demand_Management_Checks.xlsx"
    st.download_button("Download Demand Management workbook", data, file_name,
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")
    output = st.text_input("Optional local output path", value=str(Path(folder) / file_name))
    if st.button("Write workbook to local path", width="stretch"):
        try:
            path = write_demand_excel(model, output); st.success(f"Written: {path}")
        except Exception as exc: st.error(f"Could not write workbook: {exc}")
    command = f'python demand_engine.py --folder "{folder}" --output "{file_name}"'
    if simulation.strip(): command += f' --simulation "{simulation.strip()}"'
    st.code(command, language="powershell")
