"""
inv_balance_report.py v2.5.0 - Weekly inventory balance by Resource x Plant x Category x Form x Sub-class/Container
================================================================================================

Columns: Resource | Plant | Category | Form | Sub-class / Container | Measure | <ESP weeks...> | Horizon
Measures: Opening inventory, Production, Receipts (transfer/purchase/sched.), Consumption (to next stage),
          Dispatch - customer/zone, Dispatch - inter-plant transfer, Closing inventory, In-transit to plant (closing)

Two ways to run (in VS Code terminal, ibp environment):
  1) Excel only:
       C:\\Users\\rakesh.ranjan\\venvs\\ibp\\Scripts\\python.exe inv_balance_report.py "<ESP_RUN_FILES folder>"
     -> writes Inventory_Balance_Weekly_Resource.xlsx into that folder (grouped rows, +/- drill-down)
  2) On screen (filters + download button):
       C:\\Users\\rakesh.ranjan\\venvs\\ibp\\Scripts\\python.exe -m streamlit run inv_balance_report.py --server.port 8503

Sources (latest file per prefix, CSV or XLSX, kg -> MT):
  ESP_INVENTORY          opening (AVAILDATE <= first week) + scheduled receipts (later dates)
  ESP_RESLOADDETAIL      authoritative resource production: Output Qty / WhenLoaded;
                         item-location assigned to resource with largest total Output Qty
  ESP_PLANORDER          fallback production only when ESP_RESLOADDETAIL is missing
  FG_Production_Plan     FG production fallback if ESP_PLANORDER is missing
  ESP_DEPDMDSTATIC       consumption (dependent demand) - optional
  ESP_PLANARRIV          inter-plant receipts / dispatches, shipments to V_ demand zones, in-transit
  ESP_PURCHORDER         purchase receipts
  ESP_INDDMDVIEW         customer dispatch at plants (SCHEDQTY on SCHEDDATE) + item -> BU / container / sub-class
  ESP_PRODUCTIONSTEP     item/method -> resource (machine centre)
  RESOURCE_UTILIZATION   week buckets (falls back to 1st/8th/15th/22nd buckets)
"""
import os
import re
import io
import sys
import glob
import time
import datetime as dt

import numpy as np
import pandas as pd

DEFAULT_FOLDER = r"C:\Users\rakesh.ranjan\python projects\by ibp tests v2.5\ESP_RUN_FILES"
APP_OUTPUT_DIRNAME = "APP_OUTPUT_FILES"
UNIT = 1000.0
MEASURES = ["Opening inventory", "Production", "Receipts (transfer/purchase/sched.)", "Consumption (to next stage)",
            "Dispatch - customer/zone", "Dispatch - inter-plant transfer", "Closing inventory", "In-transit to plant (closing)"]
FLOW = MEASURES[1:6]
CAT_ORDER = {"RM": 1, "WIP": 2, "FG": 3, "Scrap": 4, "Other": 9}
FORM_ORDER = ["Hot metal", "RI", "CC", "HR", "CR", "CF", "FS", "FR", "AN", "FAN", "BL", "EX", "AG", "SECTION", "FG", "Scrap"]
NO_RES = "(no resource / purchased)"


def output_folder(folder):
    path = os.path.join(folder, APP_OUTPUT_DIRNAME); os.makedirs(path, exist_ok=True); return path


def timestamped_path(folder, stem, suffix=".xlsx"):
    stamp = dt.datetime.now().strftime("%d%b_%H%M")
    path = os.path.join(output_folder(folder), f"{stem}_{stamp}{suffix}")
    if not os.path.exists(path): return path
    return os.path.join(output_folder(folder), f"{stem}_{stamp}_{dt.datetime.now():%S}{suffix}")


def log(msg):
    print(f"[{dt.datetime.now():%H:%M:%S}] {msg}", flush=True)


# ------------------------------------------------------------------ loading helpers
def latest(folder, prefix, exclude=()):
    fs = []
    for f in glob.glob(os.path.join(folder, "*")):
        n = os.path.basename(f).upper()
        if n.startswith("~$") or not n.startswith(prefix.upper()) or os.path.getsize(f) < 10:
            continue
        if not n.endswith((".CSV", ".XLSX", ".XLSM", ".XLS")):
            continue
        if any(n.startswith(e.upper()) for e in exclude):
            continue
        fs.append(f)
    return max(fs, key=os.path.getmtime) if fs else None


def read(folder, prefix, exclude=()):
    p = latest(folder, prefix, exclude)
    if p is None:
        log(f"  (missing) {prefix}")
        return None
    t0 = time.time()
    if p.lower().endswith(".csv"):
        df = pd.read_csv(p, dtype=str, keep_default_na=False, encoding="utf-8-sig", low_memory=False)
    else:
        df = pd.read_excel(p, dtype=str).fillna("")
    df.columns = [str(c).replace("*", "").strip().upper() for c in df.columns]
    log(f"  loaded {os.path.basename(p)}: {len(df):,} rows ({time.time() - t0:.1f}s)")
    return df


def col(df, *names):
    for n in names:
        if n in df.columns:
            return df[n].astype(str).str.strip()
    return None


def num(s):
    return pd.to_numeric(s.astype(str).str.replace(",", "", regex=False).str.strip(), errors="coerce").fillna(0.0)


def dte(s):
    s = s.astype(str).str.strip()
    out = pd.to_datetime(s, format="%d-%b-%y %I:%M %p", errors="coerce")
    for fmt in ("%d-%b-%y", "%d-%b-%Y %I:%M %p", "%d-%b-%Y"):
        m = out.isna() & s.ne("")
        if not m.any():
            break
        out[m] = pd.to_datetime(s[m], format=fmt, errors="coerce")
    m = out.isna() & s.ne("")
    if m.any():
        out[m] = pd.to_datetime(s[m], errors="coerce", dayfirst=True, format="mixed")
    out = out.where((out > pd.Timestamp("1990-01-01")) & (out < pd.Timestamp("2100-01-01")))
    return out.dt.normalize()


def bucket(d):
    day = d.dt.day
    b = np.select([day >= 22, day >= 15, day >= 8], [22, 15, 8], 1)
    return pd.to_datetime(pd.DataFrame({"year": d.dt.year, "month": d.dt.month, "day": b}), errors="coerce")


def classify(item):
    s = str(item).upper().strip()
    p = s.split("_")
    p1 = p[1] if len(p) > 1 else ""
    p2 = p[2] if len(p) > 2 else ""
    if s.startswith("HOT_METAL"):
        return "RM", "Hot metal", "Hot metal"
    if "SCRAP" in s:
        return "Scrap", "Scrap", "Scrap"
    if s[:2] in ("R_", "F_") and p1 in ("RI", "CC", "HR", "CR", "CF", "FR", "AN", "FAN"):
        return "WIP", p1, p1
    if s.startswith("PBL"):
        return "WIP", "BL", "BL"
    if s.startswith("R_FG_") or s.startswith("F_FG_"):
        return "FG", "FG", p2 or "(blank)"
    if s.startswith("P_"):
        return "FG", "FG", {"ALLOYINGOT": "Alloy Ingot", "WR": "Wirerod", "CASTBAR": "Wirerod",
                            "HOT": "Hot Metal", "SOW": "Ingot", "PIGS": "Ingot", "RSI": "Ingot"}.get(p1, "Ingot")
    if re.match(r"^\d+$", s):
        return "WIP", "SECTION", "SECTION"
    m = re.search(r"-(EXT|AGE|ANO|PC|FAB)$", s)
    if m:
        return "WIP", {"EXT": "EX", "AGE": "AG"}.get(m.group(1), m.group(1)), {"EXT": "EX", "AGE": "AG"}.get(m.group(1), m.group(1))
    if s[:1] in ("E", "I"):
        return "FG", "FG", "Container (n/a)"
    return "Other", "Other", "Other"


# ------------------------------------------------------------------ build
def build(folder):
    T0 = time.time()
    log(f"Folder: {folder}")

    # ---- item master: BU / container / sub-class from demand view, fallback item-code rules
    imap = {}
    dv = read(folder, "ESP_INDDMDVIEW")
    if dv is not None:
        it, bu, cont, sub = col(dv, "ITEM"), col(dv, "U_PRODUCTLVL_1"), col(dv, "U_CONTAINER"), col(dv, "U_SUB_CLASS")
        if it is not None:
            m = pd.DataFrame({"I": it, "B": bu if bu is not None else "", "C": cont if cont is not None else "",
                              "S": sub if sub is not None else ""}).drop_duplicates("I")
            for i, b, c, s in m.itertuples(index=False):
                c = re.sub(r"\.0$", "", c)
                grp = f"Container {c}" if b == "E" and c else (s if s else None)
                if grp:
                    imap[i] = grp

    cls_cache = {}

    def cls(item):
        if item not in cls_cache:
            c, f, g = classify(item)
            if c == "FG" and item in imap:
                g = imap[item]
            cls_cache[item] = (c, f, g)
        return cls_cache[item]

    # ---- authoritative resource map from RESLOADDETAIL.  All inventory and
    # flow rows for an item-location must be booked to the same resource row so
    # Opening + flows = Closing.  Production-step routing is only a fallback
    # when RESLOADDETAIL itself is absent.
    meth_res, itemloc_res = {}, {}
    rl = read(folder, "ESP_RESLOADDETAIL")
    resload_prod = None
    if rl is not None and col(rl, "RES") is not None:
        resload_prod = pd.DataFrame({
            "Item": col(rl, "ITEM"), "Plant": col(rl, "LOC"), "Res": col(rl, "RES"),
            "D": dte(col(rl, "WHENLOADED")), "MT": num(col(rl, "OUTPUT QTY", "OUTPUT_QTY")) / UNIT,
            "SourceForm": col(rl, "U_PRODUCT_FORM") if col(rl, "U_PRODUCT_FORM") is not None else "",
        })
        resload_prod = resload_prod[resload_prod.Res.ne("") & ~resload_prod.Res.str.upper().str.contains("DUMMY")]
        totals = resload_prod.groupby(["Item", "Plant", "Res"], as_index=False).MT.sum()
        winners = totals.loc[totals.groupby(["Item", "Plant"]).MT.idxmax()]
        itemloc_res = winners.set_index(["Item", "Plant"]).Res.to_dict()
    else:
        ps = read(folder, "ESP_PRODUCTIONSTEP")
        if ps is not None and col(ps, "RES") is not None:
            st = pd.DataFrame({"I": col(ps, "ITEM"), "L": col(ps, "LOC"), "R": col(ps, "RES"),
                               "M": col(ps, "PRODUCTIONMETHOD") if col(ps, "PRODUCTIONMETHOD") is not None else "",
                               "N": num(col(ps, "STEPNUM")) if col(ps, "STEPNUM") is not None else 0})
            st = st[st.R.ne("") & ~st.R.str.upper().str.contains("DUMMY")]
            last = st.sort_values("N").groupby("M").R.last()
            meth_res = last.to_dict()
            itemloc_res = st.groupby(["I", "L"]).R.agg(lambda s: s.value_counts().index[0]).to_dict()

    def res_of(item, loc, method=""):
        if method and method in meth_res:
            return meth_res[method]
        return itemloc_res.get((item, loc), NO_RES)

    # ---- candidate ESP week buckets (the first displayed bucket is corrected
    # below to the authoritative ESP_INVENTORY AVAILDATE plan start).
    ru = read(folder, "RESOURCE_UTILIZATION")
    weeks = []
    if ru is not None:
        for c in ru.columns:
            d = pd.to_datetime(c, format="%d-%b-%y %I:%M %p", errors="coerce")
            if pd.isna(d):
                d = pd.to_datetime(c, format="%d-%b-%y", errors="coerce")
            if pd.notna(d):
                weeks.append(d.normalize())
    rows = []

    def add(df, measure, date_col="D"):
        if df is None or not len(df):
            return
        df = df[df.MT != 0].copy()
        df["W"] = bucket(df[date_col])
        rows.append(df[["Plant", "Item", "Res", "W", "MT"]].assign(Measure=measure))

    # ---- inventory
    iv = read(folder, "ESP_INVENTORY")
    inv = None
    if iv is not None:
        inv = pd.DataFrame({"Item": col(iv, "ITEM"), "Plant": col(iv, "LOC"), "D": dte(col(iv, "AVAILDATE")),
                            "MT": num(col(iv, "QTY")) / UNIT})
        inv["Res"] = [res_of(i, l) for i, l in zip(inv.Item, inv.Plant)]

    # ESP_INVENTORY AVAILDATE is the authoritative start of the weekly horizon.
    # This prevents old customer-order/flow dates from creating prior-year
    # buckets and the artificial spikes reported in v2.4.0.
    plan_start = inv.D.dropna().min() if inv is not None and inv.D.notna().any() else None

    # ---- production: Output Qty from RESLOADDETAIL is authoritative.
    if resload_prod is not None:
        p = resload_prod.copy()
        # If an item appears on several resources, retain its rows only on the
        # winner (largest horizon Output Qty), matching the balance resource map.
        p["Winner"] = [itemloc_res.get((i, l), NO_RES) for i, l in zip(p.Item, p.Plant)]
        p = p[p.Res.eq(p.Winner)].drop(columns="Winner")
        add(p, "Production")
    else:
        po = read(folder, "ESP_PLANORDER")
        if po is not None:
            dcol = col(po, "NEEDDATE", "SCHEDDATE", "STARTDATE")
            p = pd.DataFrame({"Item": col(po, "ITEM"), "Plant": col(po, "LOC"), "D": dte(dcol), "MT": num(col(po, "QTY")) / UNIT,
                              "M": col(po, "PRODUCTIONMETHOD") if col(po, "PRODUCTIONMETHOD") is not None else ""})
            p["Res"] = [res_of(i, l, m) for i, l, m in zip(p.Item, p.Plant, p.M)]
            add(p, "Production")
        else:
            fg = read(folder, "FG_PRODUCTION_PLAN")
            if fg is not None:
                bu, cont, sub = col(fg, "U_PRODUCTLVL_1"), col(fg, "CONTAINER", "U_CONTAINER"), col(fg, "U_SUB_CLASS")
                grp = np.where(bu == "E", "Container " + cont.str.replace(r"\.0$", "", regex=True), sub.replace("", "(blank)"))
                p = pd.DataFrame({"Item": ["@FG@" + g for g in grp], "Plant": col(fg, "PLANT"), "D": dte(col(fg, "SUPPLYAVAILDATE")),
                                  "MT": num(col(fg, "SUPPLYPEGQTY")) / UNIT, "Res": NO_RES})
                for g in set(grp): cls_cache["@FG@" + g] = ("FG", "FG", g)
                add(p, "Production")

    # ---- consumption (dependent demand)
    dd = read(folder, "ESP_DEPDMDSTATIC")
    if dd is not None and col(dd, "ITEM") is not None:
        c = pd.DataFrame({"Item": col(dd, "ITEM"), "Plant": col(dd, "LOC"),
                          "D": dte(col(dd, "STARTDATE", "NEEDDATE", "SCHEDDATE")),
                          "MT": num(col(dd, "QTY", "DMDQTY", "TOTDMD")) / UNIT})
        c["Res"] = [res_of(i, l) for i, l in zip(c.Item, c.Plant)]
        add(c, "Consumption (to next stage)")

    # ---- transfers / zone shipments / in-transit
    pa = read(folder, "ESP_PLANARRIV")
    intransit = None
    if pa is not None:
        a = pd.DataFrame({"Item": col(pa, "ITEM"), "Src": col(pa, "SOURCE"), "Dst": col(pa, "DEST"),
                          "S": dte(col(pa, "SCHEDSHIPDATE")), "A": dte(col(pa, "SCHEDARRIVDATE")), "MT": num(col(pa, "QTY")) / UNIT})
        zone = a.Dst.str.startswith("V_")
        out = a.rename(columns={"Src": "Plant", "S": "D"}).copy()
        out["Res"] = [res_of(i, l) for i, l in zip(out.Item, out.Plant)]
        add(out[zone.values], "Dispatch - customer/zone")
        add(out[~zone.values], "Dispatch - inter-plant transfer")
        rec = a[~zone].rename(columns={"Dst": "Plant", "A": "D"}).copy()
        rec["Res"] = [res_of(i, l) for i, l in zip(rec.Item, rec.Plant)]
        add(rec, "Receipts (transfer/purchase/sched.)")
        intransit = a[~zone].copy()
        intransit["Res"] = [res_of(i, l) for i, l in zip(intransit.Item, intransit.Dst)]

    # ---- purchase orders
    pu = read(folder, "ESP_PURCHORDER")
    if pu is not None:
        q = pd.DataFrame({"Item": col(pu, "ITEM"), "Plant": col(pu, "LOC"), "D": dte(col(pu, "DUEDATE")),
                          "MT": num(col(pu, "QTY")) / UNIT, "Res": NO_RES})
        add(q, "Receipts (transfer/purchase/sched.)")

    # ---- customer dispatch at plants (demand lines scheduled at plant locations)
    if dv is not None:
        d = pd.DataFrame({"Item": col(dv, "ITEM"), "Plant": col(dv, "LOC"), "D": dte(col(dv, "SCHEDDATE")),
                          "MT": num(col(dv, "SCHEDQTY")) / UNIT})
        d = d[~d.Plant.str.startswith("V_")]
        d["Res"] = [res_of(i, l) for i, l in zip(d.Item, d.Plant)]
        add(d, "Dispatch - customer/zone")

    # ---- week list
    if plan_start is not None:
        weeks = [w for w in weeks if pd.Timestamp(w) >= pd.Timestamp(plan_start)]
        if pd.Timestamp(plan_start) not in weeks:
            weeks.append(pd.Timestamp(plan_start))
    if not weeks:
        allw = pd.concat([r.W for r in rows]).dropna() if rows else pd.Series([], dtype="datetime64[ns]")
        first = pd.Timestamp(plan_start) if plan_start is not None else (allw.min() if len(allw) else pd.Timestamp.today().normalize())
        last = first + pd.DateOffset(months=4)
        d, weeks = first, []
        while d <= last:
            weeks.append(d)
            d = (d + pd.offsets.MonthBegin(1)) if d.day == 22 else d + pd.Timedelta(days=7)
    weeks = sorted(set(pd.Timestamp(w) for w in weeks))
    first, lastw = weeks[0], weeks[-1]

    # ---- opening & scheduled receipts
    if inv is not None:
        op = inv[inv.D.isna() | (inv.D <= first)].copy()
        op["W"] = first
        rows.append(op[["Plant", "Item", "Res", "W", "MT"]].assign(Measure="Opening inventory"))
        sc = inv[inv.D > first].copy()
        add(sc, "Receipts (transfer/purchase/sched.)")

    fact = pd.concat(rows, ignore_index=True) if rows else pd.DataFrame(columns=["Plant", "Item", "Res", "W", "MT", "Measure"])
    fact.loc[fact.W < first, "W"] = first                          # past-due flows land in first week
    fact = fact[fact.W.isin(weeks) & fact.Plant.ne("") & ~fact.Plant.str.startswith("V_")]
    c3 = [cls(i) for i in fact.Item]
    fact["Category"], fact["Form"], fact["Group"] = [x[0] for x in c3], [x[1] for x in c3], [x[2] for x in c3]
    fact = fact.rename(columns={"Res": "Resource"})
    K = ["Resource", "Plant", "Category", "Form", "Group"]
    flows = fact.groupby(K + ["Measure", "W"], as_index=False).MT.sum()

    # ---- closing (running balance per key)
    sign = np.where(flows.Measure.isin(["Consumption (to next stage)", "Dispatch - customer/zone",
                                        "Dispatch - inter-plant transfer"]), -1, 1)
    fl = flows.assign(N=flows.MT * sign)
    net = fl.groupby(K + ["W"]).N.sum().rename("N")
    mov = fl[fl.Measure != "Opening inventory"].groupby(K + ["W"]).N.sum().rename("MOV")
    keys = flows[K].drop_duplicates()
    grid = keys.merge(pd.DataFrame({"W": weeks}), how="cross")
    grid = grid.merge(net.reset_index(), on=K + ["W"], how="left").merge(mov.reset_index(), on=K + ["W"], how="left")
    grid[["N", "MOV"]] = grid[["N", "MOV"]].fillna(0)
    grid = grid.sort_values(K + ["W"])
    grid["MT"] = grid.groupby(K).N.cumsum()
    closing = grid.assign(Measure="Closing inventory")[K + ["Measure", "W", "MT"]]
    opening = grid.assign(MT=grid.MT - grid.MOV, Measure="Opening inventory")[K + ["Measure", "W", "MT"]]

    parts = [flows[flows.Measure != "Opening inventory"], opening, closing]
    if intransit is not None and len(intransit):
        it = []
        for r in intransit.itertuples(index=False):
            if pd.isna(r.S) or pd.isna(r.A):
                continue
            for w in weeks:
                if r.S <= w + pd.Timedelta(days=6) and r.A > w + pd.Timedelta(days=6):
                    it.append((r.Res, r.Dst, r.Item, w, r.MT))
        if it:
            itd = pd.DataFrame(it, columns=["Resource", "Plant", "Item", "W", "MT"])
            itd = itd[~itd.Plant.str.startswith("V_")]
            c3 = [cls(i) for i in itd.Item]
            itd["Category"], itd["Form"], itd["Group"] = [x[0] for x in c3], [x[1] for x in c3], [x[2] for x in c3]
            parts.append(itd.groupby(K + ["W"], as_index=False).MT.sum().assign(Measure="In-transit to plant (closing)"))
    long = pd.concat(parts, ignore_index=True)

    wl = [w.strftime("%d-%b-%y") for w in weeks]
    wide = long.pivot_table(index=K + ["Measure"], columns="W", values="MT", aggfunc="sum", fill_value=0.0)
    wide = wide.reindex(columns=weeks, fill_value=0.0)
    wide.columns = wl
    wide = wide.reset_index()
    # make sure every key has all 8 measures
    full = keys.merge(pd.DataFrame({"Measure": MEASURES}), how="cross")
    wide = full.merge(wide, on=K + ["Measure"], how="left").fillna({c: 0.0 for c in wl})
    wide["Horizon"] = np.where(wide.Measure == "Opening inventory", wide[wl[0]],
                       np.where(wide.Measure.isin(["Closing inventory", "In-transit to plant (closing)"]),
                                wide[wl[-1]], wide[wl].sum(axis=1)))
    wide = wide.rename(columns={"Group": "Sub-class / Container"})
    log(f"Built {len(wide):,} rows, {len(weeks)} weeks in {time.time() - T0:.0f}s")
    return wide, wl


def ordered(wide, wl):
    """Plant -> Resource -> Category -> Form/Sub-class detail (outline levels 1/2/3/4)."""
    value_cols = wl + ["Horizon"]
    cols = ["Plant", "Resource", "Category", "Form", "Sub-class / Container", "Measure"] + value_cols
    parts = []

    def total_rows(frame, plant, resource, category, form, group, level):
        total = frame.groupby("Measure", as_index=False)[value_cols].sum()
        total.insert(0, "Sub-class / Container", group)
        total.insert(0, "Form", form)
        total.insert(0, "Category", category)
        total.insert(0, "Resource", resource)
        total.insert(0, "Plant", plant)
        total["_lvl"] = level
        total["_m"] = total.Measure.map(MEASURES.index)
        return total.sort_values("_m")[cols + ["_lvl"]]

    for plant in sorted(wide.Plant.dropna().astype(str).unique()):
        p = wide[wide.Plant.astype(str).eq(plant)].copy()
        parts.append(total_rows(p, plant, "All", "All", "All", "Plant total", 1))

        resources = sorted(p.Resource.dropna().astype(str).unique(), key=lambda x: (x == NO_RES, x))
        for resource in resources:
            r = p[p.Resource.astype(str).eq(resource)].copy()
            parts.append(total_rows(r, plant, resource, "All", "All", "Resource total", 2))

            categories = sorted(r.Category.dropna().astype(str).unique(), key=lambda x: (CAT_ORDER.get(x, 9), x))
            for category in categories:
                c = r[r.Category.astype(str).eq(category)].copy()
                parts.append(total_rows(c, plant, resource, category, "All", f"{category} total", 3))

                c["_lvl"] = 4
                c["_f"] = c.Form.map(lambda f: FORM_ORDER.index(f) if f in FORM_ORDER else 50)
                c["_m"] = c.Measure.map(MEASURES.index)
                c = c.sort_values(["_f", "Form", "Sub-class / Container", "_m"])
                parts.append(c[cols + ["_lvl"]])

    return pd.concat(parts, ignore_index=True) if parts else pd.DataFrame(columns=cols + ["_lvl"])


def excel_bytes(wide, wl, folder=""):
    df = ordered(wide, wl)
    buf = io.BytesIO()
    import xlsxwriter
    wb = xlsxwriter.Workbook(buf, {"nan_inf_to_errors": True})
    ws = wb.add_worksheet("Inventory_Balance_Weekly")
    f_t = wb.add_format({"bold": True, "font_size": 14, "font_color": "#1F2A36"})
    f_s = wb.add_format({"italic": True, "font_color": "#5B6B7A"})
    f_h = wb.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#1F2A36", "align": "center", "border": 1})
    f_hz = wb.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#E8702A", "align": "center", "border": 1})
    tints = {"RM": "#FDEBD0", "WIP": "#FCF3CF", "FG": "#D5F5E3", "Scrap": "#FADBD8", "Other": "#EEEEEE"}
    f_tot = {k: wb.add_format({"bold": True, "bg_color": v}) for k, v in tints.items()}
    f_totn = {k: wb.add_format({"bold": True, "bg_color": v, "num_format": '#,##0.0;[Red]-#,##0.0;"-"'}) for k, v in tints.items()}
    f_txt = wb.add_format({})
    f_bold = wb.add_format({"bold": True})
    f_n = wb.add_format({"num_format": '#,##0.0;[Red]-#,##0.0;"-"'})
    f_nb = wb.add_format({"bold": True, "bg_color": "#E6EEF1", "num_format": '#,##0.0;[Red]-#,##0.0;"-"'})
    ws.write(0, 0, "Inventory balance - weekly drill-down (MT): Plant -> Resource -> Category -> Form -> Sub-class / Container", f_t)
    ws.write(1, 0, "Use the [+]/[-] buttons or outline levels 1/2/3/4 at left. Closing = Opening + Production + Receipts - Consumption - Dispatch. "
                   "Next week Opening = previous Closing. Red = negative (plan gap).", f_s)
    head = list(df.columns[:-1])
    for j, h in enumerate(head):
        ws.write(3, j, h, f_hz if h == "Horizon" else f_h)
    widths = [9, 26, 9, 10, 22, 34] + [10] * (len(wl) + 1)
    for j, w in enumerate(widths):
        ws.set_column(j, j, w)
    vals = df[head].to_numpy()
    lvls, cats, meas = df._lvl.to_numpy(), df.Category.to_numpy(), df.Measure.to_numpy()
    block_rows = {}
    for i, rec in df.iterrows():
        key = tuple(rec[c] for c in ["Plant", "Resource", "Category", "Form", "Sub-class / Container"])
        block_rows.setdefault(key, {})[rec["Measure"]] = 5 + i  # Excel 1-based row
    from xlsxwriter.utility import xl_col_to_name
    for i in range(len(df)):
        r = 4 + i
        summary = int(lvls[i]) <= 3
        c = cats[i] if cats[i] in tints else "Other"
        ws.set_row(r, None, None, {"level": int(lvls[i]), "hidden": int(lvls[i]) > 1})
        for j in range(6):
            ws.write(r, j, vals[i][j], f_tot[c] if summary else (f_bold if j == 5 and meas[i] in ("Opening inventory", "Closing inventory") else f_txt))
        bold_close = meas[i] == "Closing inventory"
        key = tuple(vals[i][j] for j in range(5)); rows_for = block_rows[key]
        for j in range(6, len(head)):
            fmt = f_totn[c] if summary else (f_nb if bold_close else f_n)
            if j < 6 + len(wl):
                week_i = j - 6
                if meas[i] == "Opening inventory" and week_i > 0:
                    formula = f"={xl_col_to_name(j-1)}{rows_for['Closing inventory']}"
                    ws.write_formula(r, j, formula, fmt)
                elif meas[i] == "Closing inventory":
                    colx = xl_col_to_name(j)
                    formula = (f"={colx}{rows_for['Opening inventory']}+{colx}{rows_for['Production']}+"
                               f"{colx}{rows_for['Receipts (transfer/purchase/sched.)']}-"
                               f"{colx}{rows_for['Consumption (to next stage)']}-"
                               f"{colx}{rows_for['Dispatch - customer/zone']}-"
                               f"{colx}{rows_for['Dispatch - inter-plant transfer']}")
                    ws.write_formula(r, j, formula, fmt)
                else:
                    ws.write_number(r, j, round(float(vals[i][j]), 3), fmt)
            else:
                first_col, last_col = xl_col_to_name(6), xl_col_to_name(5 + len(wl))
                if meas[i] == "Opening inventory": formula = f"={first_col}{r+1}"
                elif meas[i] in ("Closing inventory", "In-transit to plant (closing)"): formula = f"={last_col}{r+1}"
                else: formula = f"=SUM({first_col}{r+1}:{last_col}{r+1})"
                ws.write_formula(r, j, formula, fmt)
    ws.freeze_panes(4, 6)
    ws.autofilter(3, 0, 3 + len(df), len(head) - 1)
    ws.outline_settings(True, False, True, False)
    ws.conditional_format(4, 6, 3 + len(df), len(head) - 1,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": wb.add_format({"font_color": "#C00000"})})
    info = wb.add_worksheet("About")
    lines = ["Source folder: " + folder, "Generated: " + dt.datetime.now().strftime("%d-%b-%Y %H:%M"),
             "Units: MT (ESP kg / 1,000)", "Resource = largest ESP_RESLOADDETAIL Output Qty for item-location (ESP_PRODUCTIONSTEP fallback); purchased / no-route items: " + NO_RES,
             "Opening = ESP_INVENTORY on/before first week; later-dated inventory = scheduled receipts",
             "Production = ESP_RESLOADDETAIL Output Qty (ESP_PLANORDER / FG_Production_Plan fallback); Consumption = ESP_DEPDMDSTATIC",
             "Receipts = ESP_PLANARRIV arrivals + ESP_PURCHORDER + scheduled inventory",
             "Dispatch customer/zone = ESP_INDDMDVIEW SCHEDQTY at plants + ESP_PLANARRIV to V_ zones; inter-plant = ESP_PLANARRIV to plants",
             "In-transit = ESP_PLANARRIV shipped by week end but not yet arrived (memo, not in closing)"]
    for i, l in enumerate(lines):
        info.write(i, 0, l)
    info.set_column(0, 0, 140)
    wb.close()
    return buf.getvalue()


# ------------------------------------------------------------------ screen mode (streamlit)
def app():
    import streamlit as st
    st.set_page_config(page_title="Inventory balance by resource", layout="wide")
    folder = st.sidebar.text_input("ESP export folder", DEFAULT_FOLDER)
    if not os.path.isdir(folder):
        st.error("Folder not found")
        st.stop()
    st.sidebar.caption(f"OUTPUT_FOLDER = <ESP folder>\\{APP_OUTPUT_DIRNAME}")

    @st.cache_data(show_spinner="Building weekly balance ...")
    def get(folder, stamp):
        return build(folder)

    stamp = max((os.path.getmtime(f) for f in glob.glob(os.path.join(folder, "*"))), default=0)
    if st.sidebar.button("Reload"):
        st.cache_data.clear()
    wide, wl = get(folder, stamp)
    st.title("Inventory balance - weekly (MT) by Plant, Resource, Category, Form, Sub-class")

    def ms(label, s):
        return st.sidebar.multiselect(label, sorted(s.dropna().unique().tolist()))
    fp, fr, fc = ms("Plant", wide.Plant), ms("Resource", wide.Resource), ms("Category", wide.Category)
    ff, fg = ms("Form", wide.Form), ms("Sub-class / Container", wide["Sub-class / Container"])
    fm = st.sidebar.multiselect("Measure", MEASURES, default=MEASURES)
    d = wide
    for c, v in (("Plant", fp), ("Resource", fr), ("Category", fc), ("Form", ff), ("Sub-class / Container", fg)):
        if v:
            d = d[d[c].isin(v)]
    view = st.radio(
        "View",
        ["Detail (Form / Sub-class)", "Resource x Category totals", "Resource totals", "Plant totals"],
        horizontal=True,
    )
    show = ordered(d, wl)
    level = 4 if view.startswith("Detail") else (3 if "Category" in view else (2 if view == "Resource totals" else 1))
    show = show[show._lvl == level].drop(columns="_lvl")
    show = show[show.Measure.isin(fm)]
    st.caption(f"{len(show):,} rows · Closing = Opening + Production + Receipts - Consumption - Dispatch")
    st.dataframe(show.round(1), hide_index=True, width="stretch", height=600)
    if st.button("Prepare timestamped Excel (drill-down)"):
        data = excel_bytes(d, wl, folder); dest = timestamped_path(folder, "Inventory_Balance_Weekly_Resource")
        with open(dest, "wb") as fh: fh.write(data)
        st.session_state["inv_balance_bytes"], st.session_state["inv_balance_name"] = data, os.path.basename(dest)
        st.success(f"Saved: {dest}")
    if "inv_balance_bytes" in st.session_state:
        st.download_button("Download Excel (drill-down)", st.session_state["inv_balance_bytes"], file_name=st.session_state["inv_balance_name"])


def _in_streamlit():
    try:
        from streamlit.runtime import exists
        return exists()
    except Exception:
        return False


if __name__ == "__main__":
    if _in_streamlit():
        app()
    else:
        folder = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_FOLDER
        wide, wl = build(folder)
        dest = timestamped_path(folder, "Inventory_Balance_Weekly_Resource")
        t0 = time.time()
        with open(dest, "wb") as fh:
            fh.write(excel_bytes(wide, wl, folder))
        log(f"Saved {dest} ({time.time() - t0:.0f}s)")
