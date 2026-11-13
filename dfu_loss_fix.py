"""
dfu_loss_fix.py (v3) - BU_FCST lost in DFU -> SKU mapping
  UDT_BU_HEAD_FCST  (demand item = DMDUNIT, DMDGROUP, LOC, SBU_FCST)
     -> DFUTOSKU_FE (DMDUNIT -> ITEM, with ratio/allocation factor)
     -> ESP_DFUTOSKUFCST (supply ITEM + DFULOC + month, TOTFCST; shared item forecast split pro-rata to DFUs)
Run:
  python dfu_loss_fix.py "<ESP_RUN_FILES>" 2026-10 2027-02
Output: DFU_Loss_Fixed.xlsx in the same folder + printed reconciliation.
"""
import sys, os, glob, datetime as dt
import numpy as np
import pandas as pd

folder = sys.argv[1] if len(sys.argv) > 1 else "."
APP_OUTPUT_DIRNAME = "APP_OUTPUT_FILES"
m_from = pd.Period(sys.argv[2], "M") if len(sys.argv) > 2 else None
m_to = pd.Period(sys.argv[3], "M") if len(sys.argv) > 3 else None


def output_folder():
    path = os.path.join(folder, APP_OUTPUT_DIRNAME); os.makedirs(path, exist_ok=True); return path


def timestamped_path(stem, suffix=".xlsx"):
    now = dt.datetime.now(); stamp = now.strftime("%d%b_%H%M")
    path = os.path.join(output_folder(), f"{stem}_{stamp}{suffix}")
    if not os.path.exists(path): return path
    return os.path.join(output_folder(), f"{stem}_{stamp}_{now:%S}{suffix}")


def latest(prefix, exclude=()):
    fs = [f for f in glob.glob(os.path.join(folder, prefix + "*"))
          if f.lower().endswith((".csv", ".xlsx")) and not os.path.basename(f).startswith("~$")
          and not any(os.path.basename(f).upper().startswith(e) for e in exclude)]
    if not fs:
        sys.exit(f"Missing file starting with {prefix} in {folder}")
    return max(fs, key=os.path.getmtime)


def read(prefix, exclude=()):
    p = latest(prefix, exclude)
    df = (pd.read_csv(p, dtype=str, keep_default_na=False, encoding="utf-8-sig", low_memory=False)
          if p.lower().endswith(".csv") else pd.read_excel(p, dtype=str).fillna(""))
    df.columns = [c.replace("*", "").strip().upper() for c in df.columns]
    print(f"read {os.path.basename(p)}: {len(df):,} rows | columns: {', '.join(df.columns)}")
    return df


def pick(df, names, what, required=True):
    for n in names:
        if n in df.columns:
            return n
    if required:
        sys.exit(f"Could not find {what} column. Looked for {names}. Columns are: {list(df.columns)}")
    return None


num = lambda s: pd.to_numeric(s.astype(str).str.replace(",", "", regex=False).str.strip(), errors="coerce")
txt = lambda s: s.astype(str).str.strip()
grp = lambda s: txt(s).str.split("_").str[0]            # 20168393_PHARMA -> 20168393


def month(s):
    s = txt(s)
    d = pd.to_datetime(s, format="%d-%b-%y %I:%M %p", errors="coerce")
    for fmt in ("%d-%b-%y", "%d-%b-%Y %I:%M %p", "%d-%b-%Y"):
        m = d.isna() & s.ne("")
        if m.any():
            d[m] = pd.to_datetime(s[m], format=fmt, errors="coerce")
    m = d.isna() & s.ne("")
    if m.any():
        d[m] = pd.to_datetime(s[m], errors="coerce", dayfirst=True, format="mixed")
    return d.dt.to_period("M")


bu = read("UDT_BU_HEAD_FCST")
fe = read("DFUTOSKU_FE")
sk = read("ESP_DFUTOSKUFCST", exclude=("ESP_DFUTOSKUFCST_BU",))

# ---------- DFUTOSKU_FE mapping: DMDUNIT (+DMDGROUP) + DFULOC -> ITEM (+SKULOC), ratio
c_du = pick(fe, ["DMDUNIT"], "demand unit")
c_it = pick(fe, ["ITEM"], "supply item")
c_dl = pick(fe, ["DFULOC", "LOC"], "DFU location")
c_dg = pick(fe, ["DMDGROUP"], "demand group", required=False)
c_r = pick(fe, ["ALLOCFACTOR", "ALLOC_FACTOR", "CONVFACTOR", "CONV_FACTOR", "FACTOR", "RATIO", "SPLIT", "SPLITPCT",
                "ALLOCPCT", "PCT", "PERCENT", "DFUTOSKUFACTOR"], "ratio", required=False)
mp = pd.DataFrame({"DFU": txt(fe[c_du]), "ITEM": txt(fe[c_it]), "LOC": txt(fe[c_dl]),
                   "G": grp(fe[c_dg]) if c_dg else "*",
                   "RATIO": num(fe[c_r]) if c_r else 1.0})
mp["RATIO"] = mp.RATIO.fillna(1.0)
if c_r and mp.RATIO.max() > 1.5:          # percentages -> fraction
    mp["RATIO"] = mp.RATIO / 100.0
mp = mp.drop_duplicates()
print(f"\nMapping: {c_du} -> {c_it} at {c_dl}"
      + (f", group {c_dg}" if c_dg else "") + (f", ratio {c_r}" if c_r else ", ratio = 1 (no ratio column found)"))

# ---------- BU forecast (demand side)
b = pd.DataFrame({"DMDGROUP": txt(bu["DMDGROUP"]), "G": grp(bu["DMDGROUP"]), "DFU": txt(bu["DMDUNIT"]),
                  "LOC": txt(bu["LOC"]), "MONTH": month(bu["STARTDATE"]), "BU_FCST_RAW": num(bu["SBU_FCST"]).fillna(0)})
# ---------- SKU forecast (supply side)
s = pd.DataFrame({"ITEM": txt(sk["ITEM"]), "LOC": txt(sk["DFULOC"]),
                  "MONTH": month(sk["STARTDATE"]), "BU": txt(sk["U_PRODUCTLVL_1"]) if "U_PRODUCTLVL_1" in sk.columns else "",
                  "SKU_MT": num(sk["TOTFCST"]).fillna(0) / 1000})
# Contractual units: SBU_FCST is already MT; TOTFCST is kg and is divided by
# 1,000 above.  The magnitude check is diagnostic only and never changes units.
raw_total, sku_total = float(b.BU_FCST_RAW.clip(lower=0).sum()), float(s.SKU_MT.clip(lower=0).sum())
scale = 1.0; magnitude_ratio = raw_total / sku_total if sku_total > 0 else np.nan
magnitude_check = ("OK - same order of magnitude" if np.isfinite(magnitude_ratio) and .1 <= magnitude_ratio <= 10
                   else "REVIEW - totals differ by more than 10x" if np.isfinite(magnitude_ratio) else "REVIEW - SKU total unavailable")
b["BU_FCST_MT"] = b.pop("BU_FCST_RAW")
print(f"Divisors: SBU_FCST / 1 (already MT); TOTFCST / 1000 (kg -> MT) | BU/SKU={magnitude_ratio:.3f} | {magnitude_check}")
for d in (b, s):
    if m_from is not None:
        d.drop(d[d.MONTH < m_from].index, inplace=True)
    if m_to is not None:
        d.drop(d[d.MONTH > m_to].index, inplace=True)

b = b.groupby(["DMDGROUP", "G", "DFU", "LOC", "MONTH"], as_index=False).BU_FCST_MT.sum()
s = s.groupby(["ITEM", "LOC", "MONTH"], as_index=False).agg(SKU_MT=("SKU_MT", "sum"), BU=("BU", "first"))

# ---------- BU forecast -> expected SKU forecast via mapping and ratio
on = ["DFU", "LOC"] + (["G"] if c_dg else [])
x = b.merge(mp if c_dg else mp.drop(columns="G"), on=on, how="left")
x["MAPPED"] = x.ITEM.notna()
x["RATIO"] = x.RATIO.fillna(0)
x["EXP_MT"] = x.BU_FCST_MT * x.RATIO
# supply side = ITEM + DFULOC + month (shared by all DFUs mapped to that item) -> split pro-rata to expected
exp_item = x[x.MAPPED].groupby(["ITEM", "LOC", "MONTH"]).EXP_MT.sum().rename("EXP_ITEM")
x = x.merge(exp_item.reset_index(), on=["ITEM", "LOC", "MONTH"], how="left")
x = x.merge(s, on=["ITEM", "LOC", "MONTH"], how="left")
x["SKU_MT"] = x.SKU_MT.fillna(0)
x["SKU_SHARE_MT"] = np.where(x.EXP_ITEM > 0, x.SKU_MT * x.EXP_MT / x.EXP_ITEM, 0.0)
sku_any = set(zip(s.ITEM, s.LOC))
x["SKU_ANY_MONTH"] = [(i, l) in sku_any for i, l in zip(x.ITEM, x.LOC)]

# one row per DFU-month (sum over mapped items)
r = x.groupby(["DFU", "DMDGROUP", "LOC", "MONTH"], as_index=False).agg(
    BU_FCST_MT=("BU_FCST_MT", "first"), MAPPED=("MAPPED", "any"),
    ITEMS=("ITEM", lambda v: ", ".join(sorted(set(map(str, v.dropna())))[:3])),
    RATIO_SUM=("RATIO", "sum"), EXPECTED_SKU_MT=("EXP_MT", "sum"), DFUTOSKU_FCST_MT=("SKU_SHARE_MT", "sum"),
    SKU_ANY=("SKU_ANY_MONTH", "any"), BU=("BU", "first"))
r["BU"] = r.BU.fillna(r.DFU.str[:1].map({"F": "Foils", "R": "FRP", "P": "PRIMARY", "E": "E"})).fillna("")
# SKU forecast can be shared by several DFUs of the same group: compare at DFU level using ratio-weighted expectation
r["LOST_MT"] = (r.BU_FCST_MT - r.DFUTOSKU_FCST_MT).clip(lower=0)
r["REASON"] = np.select(
    [~r.MAPPED, r.LOST_MT <= 0.0005, ~r.SKU_ANY, r.DFUTOSKU_FCST_MT <= 0.0005],
    ["No DFUTOSKU_FE mapping", "OK", "Mapped, item never in DFUTOSKUFCST", "Mapped, no SKU forecast this month"],
    "Partial (SKU fcst < BU fcst)")
r.loc[r.RATIO_SUM.round(3).ne(1) & r.MAPPED & r.REASON.ne("OK"), "REASON"] += " (ratio sum != 1)"
r["BU_DIVISOR"], r["SKU_DIVISOR"] = 1, 1000
r["MAGNITUDE_RATIO"], r["MAGNITUDE_CHECK"] = magnitude_ratio, magnitude_check
r["MONTH"] = r.MONTH.astype(str)

loss = r[r.REASON != "OK"].sort_values("LOST_MT", ascending=False)
cols = ["DFU", "DMDGROUP", "LOC", "BU", "MONTH", "ITEMS", "RATIO_SUM", "BU_FCST_MT", "EXPECTED_SKU_MT", "DFUTOSKU_FCST_MT", "LOST_MT", "REASON",
        "BU_DIVISOR", "SKU_DIVISOR", "MAGNITUDE_RATIO", "MAGNITUDE_CHECK"]
summary = pd.DataFrame({"Measure": ["SBU_FCST divisor (already MT)", "TOTFCST divisor (kg -> MT)", "BU/SKU magnitude ratio", "Magnitude check",
                                    "BU_FCST_MT", "DFUTOSKU_FCST_MT (all SKU fcst in horizon)", "Lost MT", "Lost %",
                                    "DFU-month rows", "Rows with loss", "Rows with no FE mapping"],
                        "Value": [1, 1000, round(magnitude_ratio, 4) if np.isfinite(magnitude_ratio) else "n/a", magnitude_check,
                                  round(r.BU_FCST_MT.sum(), 1), round(s.SKU_MT.sum(), 1), round(loss.LOST_MT.sum(), 1),
                                  round(100 * loss.LOST_MT.sum() / max(r.BU_FCST_MT.sum(), 1e-9), 1), len(r), len(loss),
                                  int((~r.MAPPED).sum())]})
by_reason = loss.groupby("REASON", as_index=False).agg(Rows=("LOST_MT", "size"), Lost_MT=("LOST_MT", "sum"))
by_bu = loss.groupby(["BU", "MONTH"], as_index=False).LOST_MT.sum()
print("\n" + summary.to_string(index=False) + "\n\n" + by_reason.round(1).to_string(index=False))
chk = r[r.DFU.eq("F_AA_AA8021_44.1-60_940-1005")].head(6)
if len(chk):
    print("\nCheck F_AA_AA8021_44.1-60_940-1005:\n" + chk[cols].round(1).to_string(index=False))

dest = timestamped_path("DFU_Loss_Fixed")
with pd.ExcelWriter(dest, engine="xlsxwriter", engine_kwargs={"options": {"nan_inf_to_errors": True}}) as xw:
    summary.to_excel(xw, sheet_name="Summary", index=False)
    by_reason.to_excel(xw, sheet_name="By_Reason", index=False)
    by_bu.to_excel(xw, sheet_name="By_BU_Month", index=False)
    loss[cols].head(50000).to_excel(xw, sheet_name="DFU_Loss", index=False)
    r[cols].head(200000).to_excel(xw, sheet_name="All_DFU_Months", index=False)
    for ws in xw.sheets.values():
        ws.set_column(0, 10, 18)
print(f"\nSaved: {dest}")
