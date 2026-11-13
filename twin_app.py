"""Full Streamlit front end for the ESP Network Digital Twin.

Run from the VS Code terminal:
    streamlit run twin_app.py

The app is intentionally a thin presentation layer: twin_engine.py builds every
fact and diagnostic table used here and by the Excel/CLI output.
"""
from __future__ import annotations

import os
from pathlib import Path

import altair as alt
import numpy as np
import pandas as pd
import streamlit as st

from twin_engine import (APP_OUTPUT_DIRNAME, APP_VERSION, TwinModel, build_twin, discover_simulations, excel_bytes,
                         network_animation_gif, output_folder, timestamped_output_path, write_excel, write_findings)

DEFAULT_FOLDER = os.environ.get(
    "ESP_TWIN_DATA_FOLDER",
    r"C:\Users\rakesh.ranjan\OneDrive - Aditya Birla Group\python projects\by ibp tests v2.5\ESP_RUN_FILES")
DEFAULT_SIMULATION = "MONTHLY_WHATIF"

st.set_page_config(page_title="ESP Digital Twin", page_icon="🏭", layout="wide")

try:
    import pyarrow  # noqa: F401
    ARROW_OK = True
except Exception:
    # A defensive fallback for mixed local environments (for example PyArrow
    # compiled for NumPy 2 while VS Code is running NumPy 1).  The supported
    # requirements install has Arrow available and uses the native widgets.
    ARROW_OK = False


@st.cache_data(show_spinner="Loading and transforming ESP exports…", max_entries=8)
def load_model(folder: str, simulation: str, signature: tuple) -> TwinModel:
    del signature
    return build_twin(folder, simulation.strip() or None)


@st.cache_data(show_spinner="Building Excel workbook…", max_entries=4)
def make_workbook(folder: str, simulation: str, signature: tuple) -> bytes:
    return excel_bytes(load_model(folder, simulation, signature))


@st.cache_data(show_spinner="Rendering weekly network animation…", max_entries=4)
def make_network_clip(folder: str, simulation: str, signature: tuple) -> bytes:
    clip = network_animation_gif(load_model(folder, simulation, signature))
    timestamped_output_path(folder, "ESP_Weekly_Network_Animation", ".gif").write_bytes(clip)
    return clip


@st.cache_data(show_spinner=False, max_entries=4)
def simulation_names(folder: str, signature: tuple) -> list[str]:
    del signature
    return discover_simulations(folder)


def folder_signature(folder: str) -> tuple:
    root = Path(folder).expanduser()
    if not root.exists(): return ()
    return tuple(sorted((str(p.relative_to(root)), p.stat().st_size, p.stat().st_mtime_ns)
                        for p in root.rglob("*") if p.is_file() and p.suffix.lower() in {".csv", ".txt"}
                        and APP_OUTPUT_DIRNAME.casefold() not in {part.casefold() for part in p.relative_to(root).parts}))


def choices(df: pd.DataFrame, column: str) -> list[str]:
    if column not in df: return []
    return sorted(x for x in df[column].dropna().astype(str).unique().tolist() if x not in {"", "nan", "None"})


def apply_filters(df: pd.DataFrame, filters: dict[str, list[str]]) -> pd.DataFrame:
    out = df
    for column, selected in filters.items():
        target = column
        if column == "PLANT" and column not in out:
            target = "LOC" if "LOC" in out else "LOCATION" if "LOCATION" in out else column
        if selected and target in out: out = out[out[target].astype(str).isin(selected)]
    return out


def show_table(df: pd.DataFrame, key: str, height: int = 430) -> None:
    display = df.copy()
    # Arrow cannot infer a stable type for mixed object columns (for example
    # Forecast reconciliation VALUE containing text, dates and numbers).
    for column in display.select_dtypes(include=["object"]).columns:
        display[column] = display[column].map(lambda value: "" if pd.isna(value) else str(value)).astype("string")
    if ARROW_OK:
        st.dataframe(display, width="stretch", height=height, hide_index=True, key=key)
    else:
        st.caption("HTML fallback: PyArrow is not compatible with this Python environment.")
        st.markdown(display.head(1000).to_html(index=False, border=0), unsafe_allow_html=True)


def hide_empty_columns(df: pd.DataFrame) -> pd.DataFrame:
    keep = []
    for column in df.columns:
        values = df[column]
        if values.notna().any() and not values.astype(str).isin(["", "(Blank)", "None", "nan"]).all(): keep.append(column)
    return df[keep] if keep else df


def show_line_chart(data: pd.DataFrame, **kwargs) -> None:
    if ARROW_OK:
        st.line_chart(data, width="stretch", **kwargs)
    else:
        show_table(data.reset_index(), f"fallback_chart_{abs(hash(tuple(map(str, data.columns))))}", min(int(kwargs.get("height", 390)), 500))


def fmt_mt(value: float) -> str:
    return f"{value:,.1f} MT"


def fmt_pct(value: float) -> str:
    return "n/a" if pd.isna(value) else f"{value:,.1f}%"


def metric_row(model: TwinModel) -> None:
    k = model.kpis(); cols = st.columns(8)
    cols[0].metric("Order fill", fmt_pct(float(k["Order fill %"])))
    cols[1].metric("Unallocated", fmt_mt(float(k["Unallocated MT"])))
    cols[2].metric("Bottlenecks", f"{k['Bottleneck resources']:,} / {k['Resource periods']:,}")
    cols[3].metric("FG plan", fmt_mt(float(k["FG plan MT"])))
    cols[4].metric("OTIF in horizon", fmt_pct(float(k["OTIF in horizon %"])))
    cols[5].metric("Forecast lost", fmt_mt(float(k["Forecast lost MT"])))
    cols[6].metric("Mapping issues", f"{k['DFU mapping issues']:,}")
    cols[7].metric("Files", f"{k['Source files']:,}")


def waterfall(flow: pd.DataFrame) -> None:
    stages = ["BU_FCST_MT", "DFUTOSKU_FCST_MT", "NETTED_FCST_MT", "CUSTORDER_MT", "TOTAL_FCST_MT", "SCHED_QTY_MT"]
    totals = {c: float(flow.get(c, pd.Series(dtype=float)).sum()) for c in stages}
    display = [
        ("BU forecast", totals["BU_FCST_MT"]), ("After DFU→SKU", totals["DFUTOSKU_FCST_MT"]),
        ("Netted forecast", totals["NETTED_FCST_MT"]), ("+ Customer orders", totals["NETTED_FCST_MT"] + totals["CUSTORDER_MT"]),
        ("Total forecast", totals["TOTAL_FCST_MT"]), ("Scheduled", totals["SCHED_QTY_MT"]),
    ]
    rows = []
    previous = 0.0
    for order, (stage, value) in enumerate(display):
        rows.append({"ORDER": order, "STAGE": stage, "VALUE_MT": value, "START_MT": min(previous, value),
                     "END_MT": max(previous, value), "DIRECTION": "Up" if value >= previous else "Down", "CHANGE_MT": value-previous})
        previous = value
    wf = pd.DataFrame(rows)
    chart = (alt.Chart(wf).mark_bar(size=42).encode(
        x=alt.X("STAGE:N", sort=display and [x[0] for x in display], title=None),
        y=alt.Y("START_MT:Q", title="MT"), y2="END_MT:Q",
        color=alt.Color("DIRECTION:N", scale=alt.Scale(domain=["Up", "Down"], range=["#70AD47", "#C00000"]), legend=None),
        tooltip=["STAGE", alt.Tooltip("VALUE_MT:Q", format=",.1f"), alt.Tooltip("CHANGE_MT:Q", format="+,.1f")]
    ).properties(height=360))
    labels = alt.Chart(wf).mark_text(dy=-8, fontSize=11).encode(x=alt.X("STAGE:N", sort=[x[0] for x in display]), y="END_MT:Q", text=alt.Text("VALUE_MT:Q", format=",.0f"))
    if ARROW_OK:
        st.altair_chart(chart + labels, width="stretch")
    else:
        show_table(wf[["STAGE", "VALUE_MT", "CHANGE_MT"]], "waterfall_fallback", 320)


st.title("🏭 ESP Network Digital Twin")
st.caption(f"Blue Yonder ESP service, forecast, value, capacity, inventory and network diagnostics · v{APP_VERSION}")

with st.sidebar:
    st.header("Run setup")
    folder = st.text_input("ESP export folder", value=DEFAULT_FOLDER)
    simulation = st.text_input("Simulation", value=DEFAULT_SIMULATION,
                               help="Exact match in SIMULATION_NAME/SIMULATIONNAME only. OPTIONSET is always ignored as a simulation selector.")
    st.caption("Source quantities are kg → MT unless explicitly MT; UDT_BU_HEAD_FCST.SBU_FCST is already MT.")
    st.caption(f"OUTPUT_FOLDER = <ESP folder>\\{APP_OUTPUT_DIRNAME}")
    if st.button("Reload changed files", width="stretch"): st.cache_data.clear()

signature = folder_signature(folder)
if not signature:
    st.info("Enter a folder containing ESP CSV/TXT exports.")
    st.stop()
try:
    model = load_model(folder, simulation, signature)
except Exception as exc:
    st.error(f"Could not build the digital twin: {exc}")
    st.stop()
OUTPUT_FOLDER = output_folder(folder)

trend = model.tables["weekly_trend"]
filter_columns = ["BU", "PLANT", "RESOURCE", "PRODUCT_FORM", "STREAM", "SUB_CLASS"]
filter_frames = []
for source_frame in (trend, model.tables["forecast_flow"], model.tables["service_demand"]):
    frame = source_frame[[c for c in filter_columns if c in source_frame]].dropna(axis=1, how="all")
    if not frame.empty and len(frame.columns): filter_frames.append(frame.astype("string"))
filter_universe = pd.concat(filter_frames, ignore_index=True, sort=False) if filter_frames else pd.DataFrame(columns=filter_columns)
with st.sidebar:
    st.divider(); st.header("Network filters")
    filters: dict[str, list[str]] = {}
    for column, label in [("BU", "BU"), ("PLANT", "Plant / demand zone"), ("RESOURCE", "Resource"),
                          ("PRODUCT_FORM", "Product form"), ("STREAM", "Stream"), ("SUB_CLASS", "Sub-class")]:
        opts = choices(filter_universe, column)
        filters[column] = st.multiselect(label, opts, default=[], placeholder="All", key=f"global_{column}") if opts else []

metric_row(model)
if model.warnings:
    with st.expander(f"Load notes ({len(model.warnings)})"):
        for warning in model.warnings: st.warning(warning)

(tab_overview, tab_hit, tab_service, tab_forecast, tab_value, tab_capacity, tab_root, tab_orders,
 tab_supply, tab_network, tab_animation, tab_quality, tab_weekly, tab_balance, tab_dfu, tab_plants,
 tab_files, tab_export) = st.tabs([
    "Overview", "Where to hit", "Service & demand", "Forecast", "Value at risk", "Capacity & bottlenecks",
    "Root cause of misses", "Order status / OTIF", "Supply & inventory", "Network & sourcing", "Network animation",
    "Data quality", "Weekly inventory & flow", "Inventory drill-down", "DFU loss", "Plant cards", "Files & compare", "Excel export",
])

with tab_overview:
    st.subheader("Executive network view")
    c1, c2 = st.columns([1.2, 1])
    with c1:
        hit = model.tables["where_to_hit"]
        st.markdown("**Top P1 interventions**")
        show_table(hit[hit["PRIORITY"].eq("P1")].head(20), "overview_p1", 420)
    with c2:
        weekly = apply_filters(trend, filters)
        if not weekly.empty:
            weekly = weekly.groupby(["WEEK", "MEASURE"], observed=True, as_index=False)["QTY_MT"].sum().pivot(index="WEEK", columns="MEASURE", values="QTY_MT").fillna(0)
            st.markdown("**Weekly material-flow trend**"); show_line_chart(weekly, height=420, y_label="MT")
        else: st.info("No material-flow facts match the filters.")
    with st.expander("Diagnostics — file matching, simulation filtering and step errors"):
        show_table(model.tables["diagnostics"], "overview_diagnostics", 520)

with tab_hit:
    st.subheader("P1 / P2 — where to hit")
    hit = model.tables["where_to_hit"].copy(); c1, c2 = st.columns(2)
    priorities = choices(hit, "PRIORITY"); areas = choices(hit, "AREA")
    p = c1.multiselect("Priority", priorities, default=priorities, key="hit_priority")
    a = c2.multiselect("Area", areas, default=areas, key="hit_area")
    if p: hit = hit[hit["PRIORITY"].isin(p)]
    if a: hit = hit[hit["AREA"].isin(a)]
    show_table(hit, "hit_table", 620)

with tab_service:
    st.subheader("Service and demand — UDT_ORDER_FULFILLMENT")
    service = apply_filters(model.tables["service_demand"], filters)
    demand = float(service.get("DEMAND_MT", pd.Series(dtype=float)).sum()); alloc = float(service.get("ALLOCATED_MT", pd.Series(dtype=float)).sum())
    c1, c2, c3 = st.columns(3); c1.metric("Demand", fmt_mt(demand)); c2.metric("Allocated", fmt_mt(alloc)); c3.metric("Order fill", fmt_pct(alloc/demand*100 if demand else np.nan))
    if not service.empty:
        grouped = service.groupby(["BU", "PLANT", "LOCATION_TYPE"], dropna=False, observed=True, as_index=False)[["DEMAND_MT", "ALLOCATED_MT", "UNALLOCATED_MT"]].sum()
        grouped["ORDER_FILL_PCT"] = np.where(grouped["DEMAND_MT"].gt(0), grouped["ALLOCATED_MT"] / grouped["DEMAND_MT"] * 100, np.nan)
        show_table(grouped.sort_values("UNALLOCATED_MT", ascending=False), "service_summary", 340)
    show_table(service, "service_detail", 430)

with tab_forecast:
    st.subheader("Forecast flow and netting")
    st.caption("BU_FCST_MT → DFUTOSKU_FCST_MT → NETTED_FCST_MT + CUSTORDER_MT → TOTAL_FCST_MT → SCHED_QTY_MT → MET %")
    flow = apply_filters(model.tables["forecast_flow"], filters)
    bu_match = model.tables["mapping_column_matches"].query("ROLE == 'BU forecast'")
    if not bu_match.empty:
        row = bu_match.iloc[0]
        st.info(f"BU forecast source: {row['SOURCE_MODE']} · {row['FILE']} · quantity {row['QUANTITY_COLUMN']} · date {row['MONTH_COLUMN']}")
    waterfall(flow)
    total = float(flow.get("TOTAL_FCST_MT", pd.Series(dtype=float)).sum()); sched = float(flow.get("SCHED_QTY_MT", pd.Series(dtype=float)).sum())
    c1, c2, c3 = st.columns(3); c1.metric("Total forecast", fmt_mt(total)); c2.metric("Scheduled", fmt_mt(sched)); c3.metric("MET", fmt_pct(sched/total*100 if total else np.nan))
    show_table(flow, "forecast_flow_full", 470)
    with st.expander("Forecast reconciliation and unit/horizon checks", expanded=True):
        show_table(model.tables["forecast_reconciliation"], "forecast_reconciliation", 360)
    with st.expander("BU-head weekly DFU forecast"):
        show_table(model.tables["bu_head_forecast_weekly"], "bu_head_weekly", 430)

with tab_value:
    st.subheader("Value at risk — ESP_INDDMDVIEW CALCMARGIN")
    value = apply_filters(model.tables["value_at_risk"], filters)
    c1, c2, c3 = st.columns(3)
    c1.metric("Value at risk", f"{float(value.get('VALUE_AT_RISK', pd.Series(dtype=float)).sum()):,.1f}")
    c2.metric("Negative-margin scheduled", fmt_mt(float(value.get("NEGATIVE_MARGIN_SCHEDULED_MT", pd.Series(dtype=float)).sum())))
    c3.metric("Affected rows", f"{int(value.get('CALCMARGIN', pd.Series(dtype=float)).lt(0).sum()):,}")
    show_table(value[value.get("CALCMARGIN", pd.Series(dtype=float)).lt(0)].head(500), "negative_margin", 570)

with tab_capacity:
    st.subheader("Capacity and bottlenecks")
    capacity = apply_filters(model.tables["capacity"], filters)
    bottleneck = capacity[capacity.get("BOTTLENECK", pd.Series(False, index=capacity.index)).fillna(False)]
    st.caption("Resource-level status over the plan horizon: Bottleneck = ≥98% in at least half the weeks or average utilization ≥95%; Tight = average ≥85% or ≥98% in at least a quarter of weeks; Under-used = average <30%. DUMMY resources are excluded.")
    c1, c2, c3 = st.columns(3); c1.metric("Resources", f"{len(capacity):,}"); c2.metric("Bottleneck resources", f"{len(bottleneck):,}"); c3.metric("Tight resources", f"{int(capacity.get('RESOURCE_STATUS', pd.Series(dtype=str)).eq('Tight').sum()):,}")
    show_table(hide_empty_columns(bottleneck), "capacity_bottlenecks", 400); show_table(hide_empty_columns(capacity), "capacity_all", 390)

with tab_root:
    st.subheader("Root cause of misses")
    st.caption("Unallocated demand is traced through ESP_PLANARRIV to producing plants, then through ESP_PRODUCTIONMETHOD / ESP_PRODUCTIONSTEP and ESP_BOM.")
    root = apply_filters(model.tables["root_cause_misses"], filters)
    show_table(root, "root_causes", 430)
    st.markdown("**Network trace edges**"); show_table(model.tables["network_trace"], "network_trace", 400)

with tab_orders:
    st.subheader("Customer-order status and OTIF")
    orders = apply_filters(model.tables["order_status"], filters)
    horizon = orders[orders.get("ELIGIBLE_IN_HORIZON", pd.Series(False, index=orders.index)).fillna(False)]
    order_mt = float(horizon.get("ORDER_MT", pd.Series(dtype=float)).sum()); otif_mt = float(horizon.get("OTIF_WEIGHT_MT", pd.Series(dtype=float)).sum())
    st.caption("Definition: ESP_CUSTORDERSTATUS CUST lines only; exclude SHIPDATE before plan start; a line is on-time when SCHEDSTATUS is met and SCHEDDATE ≤ SHIPDATE; headline OTIF is Qty-MT weighted.")
    c1, c2, c3 = st.columns(3); c1.metric("Eligible lines", f"{len(horizon):,}"); c2.metric("Weighted OTIF", fmt_pct(otif_mt/order_mt*100 if order_mt else np.nan)); c3.metric("Pre-plan shipments excluded", f"{int(orders.get('BEFORE_PLAN_START', pd.Series(dtype=bool)).fillna(False).sum()):,}")
    show_table(horizon, "order_horizon", 560)

with tab_supply:
    st.subheader("Supply and inventory")
    c1, c2, c3 = st.columns(3)
    inv = model.tables["inventory_health"]; placeholders = model.tables["inventory_placeholders"]; overdue = model.tables["overdue_purchase_orders"]
    c1.metric("Non-moving profiles", f"{int(inv.get('NON_MOVING', pd.Series(dtype=bool)).fillna(False).sum()):,}")
    c2.metric("Build-up profiles", f"{int(inv.get('BUILD_UP', pd.Series(dtype=bool)).fillna(False).sum()):,}")
    c3.metric("999,999 placeholders", f"{len(placeholders):,}")
    st.markdown("**PROJECTED_ON_HAND non-moving and build-up**"); show_table(inv, "inventory_health", 320)
    c4, c5 = st.columns(2)
    with c4: st.markdown("**ESP_SKUEFFINVENTORYPARAM placeholders**"); show_table(placeholders, "placeholders", 300)
    with c5: st.markdown("**Overdue ESP_PURCHORDER**"); show_table(overdue, "overdue_po", 300)

with tab_network:
    st.subheader("Network and sourcing")
    issues = model.tables["sourcing_issues"]; sourcing = model.tables["sourcing"]
    zones = filter_universe.get("PLANT", pd.Series(dtype=str)).astype(str)
    zone_count = zones[zones.str.upper().str.startswith("V_")].nunique()
    c1, c2, c3 = st.columns(3); c1.metric("Sourcing lanes", f"{len(sourcing):,}"); c2.metric("P1 issues", f"{int(issues.get('SEVERITY', pd.Series(dtype=str)).eq('P1').sum()):,}"); c3.metric("Demand zones", f"{zone_count:,}")
    show_table(issues, "sourcing_issues", 350); show_table(sourcing, "sourcing_fact", 420)

with tab_animation:
    st.subheader("Weekly material-flow network animation")
    st.caption("ESP_PLANARRIV shipments move across plant and demand-zone lanes. Plant labels show FG/WIP closing; red badges show bottleneck resource counts. The same animated GIF is embedded in the Excel Architecture sheet.")
    st.image(make_network_clip(folder, simulation, signature), width="stretch")
    c1, c2 = st.columns([1.15, 1])
    with c1:
        st.markdown("**Weekly lane detail — resource drill-down**")
        lanes = model.tables["network_lanes_weekly"]
        if lanes.empty:
            st.info("No ESP_PLANARRIV lanes were found.")
        else:
            lane_weeks = sorted(pd.to_datetime(lanes["WEEK"].dropna().unique()).tolist())
            selected_week = st.select_slider("Week bucket", lane_weeks, value=lane_weeks[0], format_func=lambda x: pd.Timestamp(x).strftime("%d-%b-%y"))
            show_table(lanes[pd.to_datetime(lanes["WEEK"]).eq(pd.Timestamp(selected_week))], "animation_lanes", 360)
    with c2:
        st.markdown("**Bottleneck machines and stage**")
        cap = model.tables["capacity"].copy()
        prod = model.tables["resource_prod_weekly"]
        if not prod.empty:
            stage = (prod.sort_values("QTY_MT", ascending=False).drop_duplicates(["PLANT", "RESOURCE"])
                     [["PLANT", "RESOURCE", "PRODUCT_FORM"]].rename(columns={"PRODUCT_FORM": "STAGE"}))
            cap = cap.merge(stage, on=["PLANT", "RESOURCE"], how="left")
        show_table(cap[cap.get("BOTTLENECK", pd.Series(False, index=cap.index)).fillna(False)], "animation_bottlenecks", 410)

with tab_quality:
    st.subheader("Data-quality and master-health checks")
    st.caption("PASS rows are retained so this page is never blank when checks execute successfully.")
    show_table(model.tables["master_health"], "master_health", 380)
    st.markdown("**DFU mapping column matches**"); show_table(model.tables["mapping_column_matches"], "mapping_matches", 260)
    st.markdown("**Full file schema**"); show_table(model.tables["schema_report"], "schema_report", 430)
    st.markdown("**Execution diagnostics**"); show_table(model.tables["diagnostics"], "quality_diagnostics", 430)
    st.info("Only SIMULATION_NAME/SIMULATIONNAME can filter rows. OPTIONSET values such as SNOPWI_BATCH_STORE_RES_PROJ remain measure-set metadata.")

with tab_weekly:
    st.subheader("Weekly inventory and material flow")
    data = apply_filters(trend, filters)
    c1, c2 = st.columns([1, 2]); grain = c1.radio("Time grain", ["Weekly", "Monthly"], horizontal=True, key="flow_grain")
    group_dims = c2.multiselect("Group by", ["BU", "PLANT", "LOCATION_TYPE", "RESOURCE", "PRODUCT_FORM", "STREAM", "SUB_CLASS"], default=["PLANT"], max_selections=3, key="flow_groups")
    period = "WEEK" if grain == "Weekly" else "MONTH"
    if data.empty: st.info("No material-flow rows match the filters.")
    else:
        grouped = data.groupby([period, *group_dims, "MEASURE"], dropna=False, observed=True, as_index=False)["QTY_MT"].sum()
        chart_data = grouped.groupby([period, "MEASURE"], observed=True, as_index=False)["QTY_MT"].sum().pivot(index=period, columns="MEASURE", values="QTY_MT").fillna(0)
        show_line_chart(chart_data, height=390, y_label="MT"); show_table(grouped, "weekly_grouped", 470)
        st.download_button("Download filtered trend CSV", grouped.to_csv(index=False).encode("utf-8-sig"), "ESP_weekly_inventory_flow.csv", "text/csv", width="stretch")

with tab_balance:
    st.subheader("Inventory balance — Plant → Category → Resource detail")
    st.caption("Expand the plant and category blocks below. Closing = Opening + Production + Receipts − Consumption − customer/zone dispatch − inter-plant dispatch; next-week Opening = prior Closing. Negative values are red.")
    bal = model.tables["inventory_balance_wide"].copy()
    if bal.empty:
        st.info("Inventory drill-down was not available. Check the load notes for missing ESP source files.")
    else:
        plant_options = sorted(bal["Plant"].dropna().astype(str).unique())
        selected_plant = st.selectbox("Plant", plant_options, key="balance_plant")
        categories = sorted(bal.loc[bal["Plant"].astype(str).eq(selected_plant), "Category"].dropna().astype(str).unique())
        selected_categories = st.multiselect("Category", categories, default=categories, key="balance_category")
        show_bal = bal[bal["Plant"].astype(str).eq(selected_plant) & bal["Category"].astype(str).isin(selected_categories)]
        numeric_cols = [c for c in show_bal.columns if c not in ["Resource", "Plant", "Category", "Form", "Sub-class / Container", "Measure"]]
        total = show_bal.groupby(["Category", "Measure"], as_index=False)[numeric_cols].sum()
        st.markdown("**Category totals**"); show_table(total.round(1), "balance_totals", 300)
        st.markdown("**Resource detail (filterable, same week columns as Excel)**")
        if ARROW_OK:
            st.dataframe(show_bal.round(1).style.format({c: "{:,.1f}" for c in numeric_cols}).map(
                lambda v: "color:#C00000;font-weight:600" if isinstance(v, (int, float, np.number)) and v < 0 else ""),
                width="stretch", height=560, hide_index=True)
        else:
            styled = show_bal.round(1).style.format({c: "{:,.1f}" for c in numeric_cols}).map(
                lambda v: "color:#C00000;font-weight:600" if isinstance(v, (int, float, np.number)) and v < 0 else "")
            st.markdown(styled.to_html(), unsafe_allow_html=True)
        for category in selected_categories:
            with st.expander(f"{selected_plant} → {category} total → resource rows"):
                detail = show_bal[show_bal["Category"].astype(str).eq(category)]
                show_table(detail.round(1), f"balance_{selected_plant}_{category}", 360)

with tab_dfu:
    st.subheader("DFU forecast lost through mapping")
    loss = apply_filters(model.tables["dfu_lost_mapping"], filters)
    if not loss.empty and {"BU_DIVISOR", "SKU_DIVISOR", "MAGNITUDE_RATIO", "MAGNITUDE_CHECK"}.issubset(loss.columns):
        u = loss.iloc[0]
        st.info(f"Units used: SBU_FCST ÷ {u['BU_DIVISOR']:.0f} (already MT); TOTFCST ÷ {u['SKU_DIVISOR']:.0f} (kg → MT). "
                f"BU/SKU magnitude ratio {u['MAGNITUDE_RATIO']:.3f}: {u['MAGNITUDE_CHECK']}")
    c1, c2 = st.columns(2); c1.metric("Affected DFU-month rows", f"{len(loss):,}"); c2.metric("Lost forecast", fmt_mt(float(loss.get("LOST_MT", pd.Series(dtype=float)).sum())))
    show_table(loss, "dfu_loss_only", 500)
    st.markdown("**Columns detected in each stage**"); show_table(model.tables["mapping_column_matches"], "dfu_matches", 260)

with tab_plants:
    st.subheader("Plant and demand-zone cards")
    cards = apply_filters(model.tables["plant_cards"], filters)
    if cards.empty: st.info("No plant-card facts are available.")
    else:
        sort_candidates = [c for c in ["SUPPLY_GAP_MT", "Production", "Dispatch", "In-Transit", "Closing Inventory", "TOTAL_FCST_MT"] if c in cards]
        sort_col = st.selectbox("Rank by", sort_candidates, key="plant_rank"); ascending = st.checkbox("Ascending", key="plant_asc")
        show_table(cards.sort_values(sort_col, ascending=ascending), "plant_cards", 600)

with tab_files:
    st.subheader("Files and simulations")
    show_table(model.tables["schema_report"], "file_catalogue", 390)
    sims = simulation_names(folder, signature)
    st.markdown("**Simulation comparison**")
    if len(sims) >= 2:
        c1, c2 = st.columns(2); left = c1.selectbox("Simulation A", sims, index=0, key="sim_a"); right = c2.selectbox("Simulation B", sims, index=1, key="sim_b")
        if st.button("Compare simulations", width="stretch"):
            a = load_model(folder, left, signature).kpis(); b = load_model(folder, right, signature).kpis()
            comp = pd.DataFrame({"KPI": sorted(set(a) | set(b))}); comp[left] = comp["KPI"].map(a); comp[right] = comp["KPI"].map(b)
            comp["DELTA"] = pd.to_numeric(comp[right], errors="coerce") - pd.to_numeric(comp[left], errors="coerce")
            show_table(comp, "simulation_compare", 420)
    else: st.caption("At least two values in SIMULATION_NAME/SIMULATIONNAME are required for comparison.")

with tab_export:
    st.subheader("Excel export")
    st.write("The workbook includes the restored service, forecast, value, capacity, root-cause, OTIF, supply, sourcing and master-health analyses plus weekly flow, DFU loss, plant cards and fact sheets.")
    name = f"ESP_Network_Digital_Twin_{simulation.strip() or 'ALL'}.xlsx"
    st.code(str(OUTPUT_FOLDER), language=None)
    if st.button("Prepare downloadable workbook", width="stretch", key="prepare_excel"):
        data = make_workbook(folder, simulation, signature)
        saved = timestamped_output_path(folder, Path(name).stem, ".xlsx"); saved.write_bytes(data)
        findings = timestamped_output_path(folder, f"Digital_Twin_Findings_{simulation.strip() or 'ALL'}", ".xlsx")
        write_findings(model, findings)
        st.session_state["esp_workbook_bytes"] = data
        st.session_state["esp_workbook_name"] = saved.name
        st.session_state["esp_workbook_saved"] = str(saved)
        st.success(f"Saved workbook and findings to {OUTPUT_FOLDER}")
    if "esp_workbook_bytes" in st.session_state:
        st.download_button("Download ESP Network Digital Twin workbook", st.session_state["esp_workbook_bytes"],
                           st.session_state.get("esp_workbook_name", name),
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")
    if st.button("Write a new timestamped workbook", width="stretch"):
        try:
            output = timestamped_output_path(folder, Path(name).stem, ".xlsx")
            st.success(f"Written: {write_excel(model, output)}")
        except Exception as exc: st.error(f"Could not write workbook: {exc}")
    st.code(f'python twin_engine.py --folder "{folder}" --simulation "{simulation}"', language="powershell")
