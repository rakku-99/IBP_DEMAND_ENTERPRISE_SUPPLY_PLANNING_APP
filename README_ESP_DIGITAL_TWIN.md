# ESP Network Digital Twin v2.5.0

## Replacement files

Copy these files together into `by ibp tests v2.4`:

- `twin_engine.py`
- `twin_app.py`
- `inv_balance_report.py`
- `dfu_loss_fix.py`
- `requirements.txt`

## Run

```powershell
cd "C:\Users\rakesh.ranjan\python projects\by ibp tests v2.5"
C:\Users\rakesh.ranjan\venvs\ibp\Scripts\python.exe -m pip install -r requirements.txt
C:\Users\rakesh.ranjan\venvs\ibp\Scripts\python.exe -m streamlit run twin_app.py
```

The default input folder is:

```text
C:\Users\rakesh.ranjan\python projects\by ibp tests v2.5\ESP_RUN_FILES
```

Every generated file is saved under:

```text
<ESP_RUN_FILES>\APP_OUTPUT_FILES
```

The folder is created automatically. Names are timestamped (`_06Nov_1035`, with seconds added if needed), so an open Excel file is never overwritten. Input discovery explicitly excludes `APP_OUTPUT_FILES`.

Command-line workbook/GIF/findings run:

```powershell
C:\Users\rakesh.ranjan\venvs\ibp\Scripts\python.exe twin_engine.py --folder "C:\Users\rakesh.ranjan\python projects\by ibp tests v2.5\ESP_RUN_FILES" --simulation MONTHLY_WHATIF
```

Standalone inventory report:

```powershell
C:\Users\rakesh.ranjan\venvs\ibp\Scripts\python.exe inv_balance_report.py "C:\Users\rakesh.ranjan\python projects\by ibp tests v2.5\ESP_RUN_FILES"
```

Standalone DFU-loss reconciliation:

```powershell
C:\Users\rakesh.ranjan\venvs\ibp\Scripts\python.exe dfu_loss_fix.py "C:\Users\rakesh.ranjan\python projects\by ibp tests v2.5\ESP_RUN_FILES" 2026-10 2027-02
```

## v2.5.0 business rules

### Weekly horizon

- First bucket = earliest `ESP_INVENTORY.AVAILDATE`.
- Arbitrary transaction dates are folded into ESP 1/8/15/22 buckets.
- The resource-utilization matrix supplies the last weekly bucket.
- Historical order/ship dates can no longer create prior-year chart spikes.

### Resource inventory drill-down

- Columns: Resource, Plant, Category, Form, Sub-class / Container, Measure, ESP weeks, Horizon.
- Hierarchy: Plant total (level 1) → category total (level 2) → resource/form/sub-class detail (level 3).
- `U_PRODUCTLVL_1 = E` uses `Container n`; other products use `U_SUB_CLASS`.
- Formula: Closing = Opening + Production + Receipts − Consumption − customer/zone dispatch − inter-plant dispatch.
- Next-week Opening is a formula reference to prior-week Closing; negatives are red.
- App provides the same plant/category/resource drill-down.

### RESLOADDETAIL production

- Resource production comes from `ESP_RESLOADDETAIL.Output Qty` using `WhenLoaded`; kg is divided by 1,000.
- Both `dd/mm/yyyy hh:mm` and `dd-Mon-yy hh:mm AM` dates are accepted.
- Each item-location is assigned to the resource with the largest horizon Output Qty.
- Opening stock, receipts, dependent demand/consumption and dispatch for that item-location are booked to the same resource row.
- `ESP_PLANORDER`/`ESP_PRODUCTIONSTEP` is used only if RESLOADDETAIL is missing.
- Verified example: `M1_M / MOU / CF / F_CF_AA_AA8021_0.210-0.400_1040-1140` contributes 310.600 MT on 01-Nov-26.

### OTIF and bottlenecks

- OTIF uses `ESP_CUSTORDERSTATUS` CUST lines, excludes `SHIPDATE` before inventory plan start, requires a met/completed status, applies `SCHEDDATE <= SHIPDATE`, and weights by Qty MT.
- Resource status is one row per distinct non-DUMMY machine centre. Reference result: 44 bottlenecks / 133 resources.

### DFU loss and units

```text
UDT_BU_HEAD_FCST DMDUNIT + normalized DMDGROUP + LOC + month
    → DFUTOSKU_FE demand item to supply ITEM × allocation ratio
    → ESP_DFUTOSKUFCST supply ITEM + DFULOC + month only
```

- Both BU and mapping DMDGROUP values are normalized to the part before the first underscore.
- `SBU_FCST` is already MT: divisor = 1.
- `TOTFCST` is kg: divisor = 1,000.
- A BU/SKU magnitude ratio and status are shown in the app, detailed DFU-loss rows, reconciliation, findings and standalone output. The check never changes the fixed divisors.

### Excel and animation

The shareable workbook follows the reference layout and includes Home, Architecture, KPI_Dashboard, Network_Map, Bottleneck_Resources, Resource_Prod_Weekly/Monthly, formula-driven Inventory_Balance_Weekly, FG/material-flow sheets, Demand_vs_Supply, In_Transit, Order_Priority, plant sheets, assumptions and facts. Headers are dark, Horizon is orange, panes/filters are frozen, and every sheet links to Home.

The animated GIF shows weekly ESP_PLANARRIV shipments among RKT, MHM, BEL, SIL, MOU, AFG, HIR, TLJ and demand zones, plus FG/WIP closing, production, customer dispatch and bottleneck machine/stage detail. It appears in the app and is embedded in the Architecture sheet.
