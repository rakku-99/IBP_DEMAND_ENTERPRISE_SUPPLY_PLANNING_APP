# ESP Network Digital Twin v2.5.0 — Test Report

## Test data

The task upload directory was empty, so the exact 5-Nov v2.4 `ESP_RUN_FILES` sources were retrieved from the user's OneDrive project folder. The real-data run used 21 relevant ESP exports containing 2,956,928 rows, including the full 729,960-row SKU forecast, 587k-row BU forecast, 375,777-row demand view, 369,875-row customer-order status, 181,174 plan orders, 145,578 dependent-demand rows and 130,401 RESLOADDETAIL rows.

A bounded schema-faithful copy of those same files was used for the all-page Streamlit render test to stay within the sandbox UI-test memory limit. The full-data model and outputs were tested separately.

## Full-data reconciliations

| Check | v2.5 result | Status |
|---|---:|---|
| First weekly bucket | 22-Oct-26 (`ESP_INVENTORY.AVAILDATE`) | Pass |
| Last weekly bucket | 22-Feb-27 | Pass |
| Weekly bucket count | 17 | Pass |
| FG production plan | 483,420.57 MT | Pass |
| Weighted OTIF | 58.4873% | Pass |
| Bottleneck resources | 44 / 133 | Pass |
| BU divisor | 1 (`SBU_FCST` already MT) | Pass |
| SKU divisor | 1,000 (`TOTFCST` kg → MT) | Pass |
| BU total in SKU horizon | 966,200.493 MT | Pass |
| SKU total in horizon | 948,516.304 MT | Pass |
| BU/SKU magnitude ratio | 1.018644 — OK | Pass |
| Network production double-counting | FG plan remains 483,420.57 MT | Pass |

The OTIF test used 1,264 eligible CUST lines after excluding pre-plan SHIPDATE. The weighted numerator was 1,384.792 MT and denominator 2,367.680 MT, yielding 58.4872956%.

## RESLOADDETAIL resource drill-down

`ESP_RESLOADDETAIL` was selected ahead of PLANORDER/PRODUCTIONSTEP. Item-location winner assignment used the resource with the largest horizon Output Qty. The exact user example was present:

```text
Resource M1_M | Plant MOU | Form CF
Item F_CF_AA_AA8021_0.210-0.400_1040-1140
WhenLoaded 01-Nov-26 | Output Qty 310,600 kg = 310.600 MT
```

The M1_M/MOU/CF resource-form total for 01-Nov-26 was 923.149 MT, including that item. Opening inventory, receipts, dependent-demand consumption and dispatch used the same item-location winner resource. The maximum computed closing-equation difference for those M1_M/MOU/CF blocks was 0.0 MT.

## Inventory hierarchy and formulas

The generated workbook contained:

- outline levels 1/2/3 plus the worksheet default level 0;
- 17,097 formulas in `Inventory_Balance_Weekly`;
- future Opening formulas referencing the preceding Closing cell;
- Closing formulas adding Opening, Production and Receipts and subtracting Consumption and both Dispatch rows;
- Horizon formulas and negative-number conditional formatting;
- frozen panes at row 4 / column 6 and filters.

A standalone schema-faithful inventory workbook was recalculated with LibreOffice: 6,929 formulas, zero formula errors.

## Workbook and animation

The real-data workbook was generated under `APP_OUTPUT_FILES` with a timestamped name. It contained the complete reference-style sheet set: Home, Architecture, KPI_Dashboard, Network_Map, Bottleneck_Resources, Resource_Prod_Weekly/Monthly, formula-driven Inventory_Balance_Weekly, inventory/FG/material-flow summaries, Demand_vs_Supply, In_Transit, Order_Priority, DFU_Loss, OTIF, item masters, assumptions, facts, data quality and plant sheets.

ZIP-level inspection confirmed both `xl/media/image1.gif` (animated network clip) and `xl/media/image2.png` (static Network_Map frame). The GIF was also written as a separate timestamped output. Dark headers, orange Horizon, Home hyperlinks, frozen panes and filters were verified.

## Streamlit render

All 18 tabs rendered on the bounded real-schema dataset with zero Streamlit exceptions:

1. Overview
2. Where to hit
3. Service & demand
4. Forecast
5. Value at risk
6. Capacity & bottlenecks
7. Root cause of misses
8. Order status / OTIF
9. Supply & inventory
10. Network & sourcing
11. Network animation
12. Data quality
13. Weekly inventory & flow
14. Inventory drill-down
15. DFU loss
16. Plant cards
17. Files & compare
18. Excel export

The full-data Streamlit AppTest process exceeded the sandbox memory limit after building the 2.96-million-row model, so UI rendering was repeated on the bounded copy. The full-data engine build, KPI calculations, resource balance, GIF and three Excel outputs completed independently.

## Output-folder safety

- `APP_OUTPUT_FILES` was created automatically.
- Main workbook, GIF and findings used timestamped names.
- Standalone DFU loss and inventory workbooks used timestamped names.
- A sentinel CSV placed under `APP_OUTPUT_FILES` was excluded from source discovery; the source count remained 21.
- Re-running in the same minute adds seconds/counters if the target already exists.

## Compatibility

- Python compilation: passed for all four replacement scripts.
- pandas 3.1-safe keyword APIs retained.
- Streamlit 1.66 tab/dataframe paths tested.
- PyArrow/NumPy incompatibility fallback tested; all pages rendered without Arrow exceptions.
- Requirements now include Pillow for GIF generation.
