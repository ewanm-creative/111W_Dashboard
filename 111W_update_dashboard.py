"""
111W_update_dashboard.py
Reads 111W_Logistics Output.xlsx and injects a DATA block into
111W_Container_Dashboard.html between the markers:
  // DATA_START
  // DATA_END
"""

import re, json, sys
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook

HERE    = Path(__file__).parent
HTML_IN = HERE / "111W_Container_Dashboard.html"
HTML_OUT = HTML_IN

if len(sys.argv) > 1:
    XL = Path(sys.argv[1])
else:
    XL = HERE / "111W_Logistics Output.xlsx"

if not XL.exists():
    sys.exit("ERROR: Excel file not found: " + str(XL))
if not HTML_IN.exists():
    sys.exit("ERROR: HTML file not found: " + str(HTML_IN))

print("Reading " + XL.name + " ...")
wb = load_workbook(XL, data_only=True)
print("  Sheets: " + str(wb.sheetnames))

def clean(v):
    if v is None: return "---"
    if isinstance(v, datetime): return v.strftime("%-m/%-d/%Y")
    return str(v).strip() or "---"

def fmt_date(v):
    if isinstance(v, datetime): return v.strftime("%b %-d, %Y")
    if v and str(v).strip() not in ("", "None", "---"):
        try:
            d = datetime.strptime(str(v).strip(), "%Y-%m-%d %H:%M:%S")
            return d.strftime("%b %-d, %Y")
        except:
            return clean(v)
    return "---"

SH_CTN = None
for name in wb.sheetnames:
    nl = name.strip().lower()
    if "container" in nl or "schedule" in nl or "ctn" in nl:
        SH_CTN = name
        break
if not SH_CTN:
    SH_CTN = wb.sheetnames[0]

print("  Container sheet: " + SH_CTN)
ws = wb[SH_CTN]

header_row = None
headers = {}
for row in ws.iter_rows(min_row=1, max_row=20):
    vals = [str(c.value).strip().lower() for c in row if c.value]
    has_num = any(v in ("cont. #", "ctn #", "container #", "ctnr no.") for v in vals)
    has_status = any(v in ("cont. status", "status") for v in vals)
    has_units = any(v == "units" for v in vals)
    if (has_num and has_status) or (has_num and has_units):
        header_row = row[0].row
        break
if not header_row:
    header_row = 1

for cell in ws[header_row]:
    if cell.value:
        headers[str(cell.value).strip().upper()] = cell.column - 1

print("  Headers (row " + str(header_row) + "): " + str(list(headers.keys())[:10]))

def col(*names):
    for n in names:
        if n.upper() in headers:
            return headers[n.upper()]
    return None

i_num   = col("CONT. #", "CTN #", "CONTAINER #", "CTN#")
i_st    = col("CONT. STATUS", "STATUS")
i_week  = col("PICKUP WEEK", "LOAD WK", "LOAD WEEK", "WK", "WEEK")
i_load  = col("LOAD DATE", "FACTORY LOAD DATE", "LOAD")
i_ship  = col("SHIP DATE", "ETD", "DEPARTURE DATE")
i_vsl   = col("VESSEL", "VESSEL NAME")
i_port  = col("ARRIVAL DATE", "PORT ARRIVAL", "ETA PORT", "ETA NY")
i_del   = col("DELIVERY DATE", "DELIVERY", "EST. DELIVERY")
i_conf  = col("CONFIRM", "CONFIRMED?", "CONFIRMED")
i_qty   = col("UNIT Q.TY", "UNIT QTY", "QTY", "UNIT QUANTITY")
i_units = col("UNITS", "UNIT RANGE", "UNIT NUMBERS", "UNIT LIST")
i_floors = col("FLOORS", "FLOOR", "FLOOR #")

CONTAINERS = []
STATUS_MAP = {"PROJ", "LDG", "LDD", "ENR", "INPRT", "D"}

for row in ws.iter_rows(min_row=header_row+1, values_only=True):
    if not any(row): continue
    def g(i, r=row): return r[i] if i is not None and i < len(r) else None
    num_raw = g(i_num)
    if not num_raw or str(num_raw).strip() in ("", "None"): continue
    if str(num_raw).strip().lower() in ("example", "ex", "#ref!"): continue
    try: num = int(float(str(num_raw).strip()))
    except: continue
    st_raw = clean(g(i_st)).upper().strip()
    if st_raw not in STATUS_MAP: st_raw = "PROJ"
    units_raw = clean(g(i_units))
    try: qty = int(float(str(g(i_qty) or "0")))
    except: qty = 0
    conf_raw = g(i_conf)
    confirmed = bool(conf_raw and str(conf_raw).strip().lower() in ("yes","y","confirmed","true","1","x"))
    CONTAINERS.append({
        "num": num, "status": st_raw, "week": clean(g(i_week)),
        "loadDate": fmt_date(g(i_load)), "shipDate": fmt_date(g(i_ship)),
        "vessel": clean(g(i_vsl)), "portArrival": fmt_date(g(i_port)),
        "delivery": fmt_date(g(i_del)), "confirmed": confirmed,
        "unitQty": qty, "units": units_raw,
        "floors": clean(g(i_floors)),
    })

print("  -> " + str(len(CONTAINERS)) + " containers parsed")

UNIT_STATUS = {}
for sname in wb.sheetnames:
    if "unit status" in sname.lower() or "exceptions" in sname.lower():
        wsu = wb[sname]
        for row in wsu.iter_rows(min_row=2, values_only=True):
            if row and row[0] and row[1]:
                try:
                    uid = int(float(str(row[0]).strip()))
                    st = str(row[1]).strip().upper()
                    if st in STATUS_MAP:
                        UNIT_STATUS[uid] = st
                except: pass
        print("  -> Unit overrides: " + str(len(UNIT_STATUS)))
        break

TRANSIT_PERF = []
SH_DASH = None
for sname in wb.sheetnames:
    nl = sname.strip().lower()
    if "111 wall" in nl or "dashboard" in nl or "transit" in nl:
        SH_DASH = sname
        break
if SH_DASH:
    wsd = wb[SH_DASH]
    for row in wsd.iter_rows(min_row=12, max_row=16, values_only=True):
        if not row: continue
        label = clean(row[0]) if len(row) > 0 else "---"
        value = clean(row[1]) if len(row) > 1 else "---"
        if label in ("---", "", "None") and len(row) > 1:
            label = clean(row[1])
            value = clean(row[3]) if len(row) > 3 else "---"
        if label not in ("---", "", "None"):
            TRANSIT_PERF.append({"label": label, "value": value})
    print("  -> Transit perf: " + str(len(TRANSIT_PERF)) + " rows")

BUILDING_MATRIX = []
SH_BM = None
for sname in wb.sheetnames:
    nl = sname.strip().lower()
    if "building matrix" in nl and "(2)" not in nl and "condensed" not in nl:
        SH_BM = sname
        break

if SH_BM:
    wsbm = wb[SH_BM]
    print("  Scanning building matrix: " + SH_BM)

    def expand_ranges(s):
        ids = []
        if not s or s in ("---", "--"): return ids
        for part in str(s).split(","):
            t = part.strip()
            if ":" in t:
                a, b = t.split(":")
                try:
                    for i in range(int(a), int(b)+1): ids.append(i)
                except: pass
            elif t and t.isdigit():
                ids.append(int(t))
        return ids

    live_status = {}
    for c in CONTAINERS:
        for uid in expand_ranges(c["units"]):
            live_status[uid] = c["status"]
    live_status.update(UNIT_STATUS)

    row_start = 7
    while row_start + 2 <= wsbm.max_row:
        row_ids    = row_start
        row_ktype  = row_start + 1
        row_status = row_start + 2
        floor_num_val = wsbm.cell(row=row_ktype, column=2).value
        if floor_num_val is None:
            row_start += 3
            continue
        try:
            floor_num = int(float(str(floor_num_val)))
        except:
            row_start += 3
            continue
        units = []
        for col_idx in range(3, 61):
            uid_val   = wsbm.cell(row=row_ids,    column=col_idx).value
            ktype_val = wsbm.cell(row=row_ktype,  column=col_idx).value
            st_val    = wsbm.cell(row=row_status, column=col_idx).value
            if uid_val is None: continue
            uid_str = str(uid_val).strip()
            if not uid_str or uid_str in ("None", "--"): continue
            ktype = str(ktype_val).strip() if ktype_val and str(ktype_val).strip() not in ("--", "None") else "--"
            try:
                uid_int = int(uid_str)
                raw_st = str(st_val).strip() if st_val and str(st_val).strip() not in ("--", "None") else "--"
                st = live_status.get(uid_int, raw_st)
            except:
                st = str(st_val).strip() if st_val else "--"
            if st.upper() not in STATUS_MAP:
                st = "--"
            units.append({"uid": uid_str, "ktype": ktype, "status": st})
        if units:
            BUILDING_MATRIX.append({"floor": floor_num, "units": units})
        row_start += 3

    print("  -> " + str(len(BUILDING_MATRIX)) + " floors parsed")
else:
    print("  ! Building matrix sheet not found")

now_str = datetime.now().strftime("%b %-d, %Y %-I:%M %p")

parts = []
parts.append("// DATA_START (auto-generated " + now_str + ")")
parts.append("const CONTAINERS = [")
for c in CONTAINERS:
    parts.append("  { num:" + str(c["num"]) + ", status:" + json.dumps(c["status"]) +
        ", week:" + json.dumps(c["week"]) + ", units:" + json.dumps(c["units"]) +
        ", unitQty:" + str(c["unitQty"]) + ", vessel:" + json.dumps(c["vessel"]) +
        ", loadDate:" + json.dumps(c["loadDate"]) + ", shipDate:" + json.dumps(c["shipDate"]) +
        ", portArrival:" + json.dumps(c["portArrival"]) + ", delivery:" + json.dumps(c["delivery"]) +
        ", confirmed:" + ("true" if c["confirmed"] else "false") +
        ", floors:" + json.dumps(c["floors"]) + " },")
parts.append("];")
parts.append("")
parts.append("const UNIT_STATUS = {")
for uid, st in UNIT_STATUS.items():
    parts.append("  " + str(uid) + ": " + json.dumps(st) + ",")
parts.append("};")
parts.append("")
parts.append("const TRANSIT_PERF = [")
for r in TRANSIT_PERF:
    parts.append("  { label:" + json.dumps(r["label"]) + ", value:" + json.dumps(r["value"]) + " },")
parts.append("];")
parts.append("")
parts.append("const LAST_UPDATED = " + json.dumps(now_str) + ";")
parts.append("")
parts.append("const BUILDING_MATRIX = [")
for f in BUILDING_MATRIX:
    uj = ", ".join("{uid:" + json.dumps(u["uid"]) + ",ktype:" + json.dumps(u["ktype"]) +
        ",status:" + json.dumps(u["status"]) + "}" for u in f["units"])
    parts.append("  { floor:" + str(f["floor"]) + ", units:[" + uj + "] },")
parts.append("];")
parts.append("// DATA_END")

data_block = "\n".join(parts)

html = HTML_IN.read_text(encoding="utf-8")
pattern = r"// DATA_START.*?// DATA_END"
if not re.search(pattern, html, flags=re.DOTALL):
    sys.exit("ERROR: DATA markers not found in HTML. Expected: // DATA_START ... // DATA_END")

html_out = re.sub(pattern, data_block, html, flags=re.DOTALL)
HTML_OUT.write_text(html_out, encoding="utf-8")
print("Dashboard updated: " + HTML_OUT.name)
print("  " + str(len(CONTAINERS)) + " containers | " + str(len(UNIT_STATUS)) +
    " unit overrides | " + str(len(TRANSIT_PERF)) + " transit rows | " +
    str(len(BUILDING_MATRIX)) + " floors in matrix")

import subprocess, os
os.chdir(HERE)
subprocess.run(["git", "add", "111W_Container_Dashboard.html"], check=False)
subprocess.run(["git", "commit", "-m", "Dashboard update " + now_str], check=False)
result = subprocess.run(["git", "push"], check=False, capture_output=True, text=True)
if result.returncode == 0:
    print("Pushed to GitHub successfully.")
else:
    print("Git push failed: " + result.stderr.strip())

if sys.platform == "win32":
    input("\nPress Enter to close...")
