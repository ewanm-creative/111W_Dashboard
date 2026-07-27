# 111 Wall Dashboard — Edge Case Test Report

> **Status: all P0 and P1 items below are fixed** (27 Jul 2026). All 50 scenarios
> now pass with no crash, no silent data loss and no JS errors.
> See "Resolution" at the end of this document.


Tested by mutating `111W_Logistics Output.xlsx` across 50+ scenarios, re-running
`111W_update_dashboard.py`, and rendering the resulting HTML headlessly to check
for crashes, silent data loss, and wrong numbers.

Baseline (unmodified file) renders clean: 31 containers, 27 floors, no JS errors.

---

## P0 — Will embarrass you in front of the client

### 1. Two whole containers are invisible right now
`Doors` (58 units) and `H1` (16 units) are rows 39–40 of the CTNR Schedule.
The parser does `int(float(num))` and skips on failure, so both are dropped
silently. Renumbering them 32/33 immediately produces 33 containers and +16
tracked units.

Consequence: the H1 hardware units (350, 351, 450, 451 … 1050, 1051) are stuck
at **Projected forever**, on every floor, regardless of what actually happens.

### 2. No floor can ever reach 100%
Setting every container to `D` and re-rendering:

| Floor | Best possible | Why |
|---|---|---|
| 3 | 91% | 3 no-kitchen units + 2 H1 units in the denominator |
| 4–10 | 97% | 2 H1 units each |
| 11 | 97% | units 1150, 1151 in no container |
| 12 | 81% | units 1248–1258 in no container |
| 13–29 | 0% | 17 empty floors rendered at "0/58" permanently |

The progress bar divides by `f.units.length` (every cell in the row) instead of
by units that actually have a kitchen. A client looking at this at project
completion sees a building that is 0–97% done.

### 3. Hard-coded "Total Containers: 37"
Line ~168 of the HTML. The script never touches it. The donut two inches away
says **31**. Both numbers are on screen simultaneously, today, and they disagree.
(`of 577` in the Units Delivered card is correct — 577 real kitchens on floors 3–12.)

### 4. A semicolon in the Units column crashes the whole update
`322:333;335:342` → `ValueError: too many values to unpack`.
The script dies **before** writing the HTML, so the site silently keeps serving
the previous day's data. Launched from `run_update.bat` by double-click, the
console window closes instantly — you get no traceback, no error, nothing.

### 5. Renaming a header silently empties the dashboard
`CONT. #` → `CTNR #` produces **0 containers**, exit code 0, and the script then
commits and pushes the empty dashboard to GitHub. Same for adding any sheet whose
name contains "container"/"schedule"/"ctn" ahead of the real one (tested with a
sheet called `Container Notes` → 0 containers, pushed).

There is no "did we parse a sane number of rows?" guard anywhere.

### 6. Inserting one row in the Building Matrix wipes the entire matrix
The parser assumes floor blocks start at row 7 and repeat every 3 rows forever.
Inserting a single row at the top → **0 floors**, both the Building Matrix tab and
the mini-matrix render empty, exit code 0, pushed to GitHub.
Inserting a column loses 9 units (the column scan is hard-coded to 3–60).

---

## P1 — Wrong numbers, no warning

| # | Trigger | Result |
|---|---|---|
| 7 | Container delivery date is in the past but status isn't `D` | Container disappears from Next Up entirely. There is no overdue/late indicator anywhere on the dashboard. |
| 8 | Status typed as `HOLD`, `CANCELLED`, `CUSTOMS`, `DELAYED` | Silently coerced to `PROJ`. A cancelled container reads as "Projected". |
| 9 | Units written `301-321` (dash) or `321:301` (reversed) | Expands to **zero** units. Container still shows Qty 20. Those 20 kitchens vanish from the matrix and the damage-report dropdown. |
| 10 | Units written `301:313,315:321 (2 short)` | Note text kills the last range — 13 units tracked instead of 20. |
| 11 | Qty typed as `20 pcs` | Parsed as 0. Units Delivered undercounts. |
| 12 | `#REF!` / `#DIV/0!` in a cell | `#REF!` in Cont. # drops the row; `#REF!` in a date prints literally in the client-facing table. |
| 13 | Unit Status override sheet used for a partial delivery | Cells turn dark in the matrix, but the **Units Delivered KPI and the donut don't move** — they only count whole containers. Two contradictory readings on one screen. |
| 14 | Same unit listed in two containers | Last one wins, no warning. |
| 15 | Duplicate container number | Two identical "CTN 5" rows, accepted. |
| 16 | Two matrix blocks with the same floor number | Renders a duplicate floor row. |
| 17 | Moving the Dashboard sheet down the tab order | Transit Performance reads garbage (`{label:"27", value:"--"}`). |
| 18 | Date typed as text `2026-08-15` | Renders **Aug 14** in New York. `new Date('2026-08-15')` parses as UTC midnight. |

Also: 13 kitchens (1150, 1151, 1248–1258) are in the Building Matrix but assigned
to no container in the source file. Combined with H1's 16, that's **29 of 577
kitchens (5%) untracked** — hidden rather than flagged.

---

## P2 — Cosmetic / hardening

- **Container table has no sort.** Reordering rows in Excel reorders the client-facing
  table. Rotating the sheet renders 16→31 then 1→15.
- **Excel formula cache.** The Building Matrix is array formulas; the script reads
  `data_only=True`. Any save by a tool that doesn't write cached values (LibreOffice,
  a script) yields a blank matrix.
- **No output escaping.** Vessel name and kitchen type go straight into HTML and into
  `title="..."` attributes. A `"` or `<` in a kitchen code produces malformed HTML.
- **MISC NOTES column is never read.** Notes typed in the master file never reach the site.
- **Header-detection mismatch.** Row detection looks for `ctnr no.`, but `col()` looks for
  `CTN#` — the two lists disagree.
- **No empty state** on the container table.
- Unit override for a non-existent unit (e.g. 9999) is accepted and does nothing.

---

## Suggested fix order

1. Guard the publish: if containers < 25 or floors < 10, print an error and **do not
   write, commit, or push**. This alone converts findings 5, 6 and 12 from silent
   corruption into a visible failure.
2. Accept non-numeric container IDs (`Doors`, `H1`, `3A`) — keep them as strings.
3. Fix the progress denominator: count only units with a kitchen type.
4. Make Total Containers dynamic; drop floors 13–29 from the matrix or label them
   "not in scope".
5. Harden `expand_ranges`: split on `,` and `;`, accept `-` and `:`, sort reversed
   ranges, strip trailing text, and **warn** on anything it can't parse instead of
   dropping it.
6. Locate the matrix by scanning for floor labels instead of assuming row 7 + 3.
7. Add an overdue flag (past delivery date, status ≠ D) to Next Up.
8. Make unit-level overrides feed the Units Delivered KPI.
9. `2>&1 | tee log.txt` + `pause` in `run_update.bat` so failures are visible.
10. Escape HTML on output; read MISC NOTES into a tooltip.

---

# Resolution — 27 Jul 2026

## What changed

**`111W_update_dashboard.py`**

- **Non-numeric container IDs are kept.** `Doors`, `H1`, `3A`, `HW2` all survive as
  their own containers. Numeric IDs sort first, named ones after, so the table
  order no longer depends on row order in the spreadsheet.
  **Container count went from 31 to 33.**
- **Sanity guard before publishing.** Compares the new parse against what is
  already published. Zero containers, zero floors, or a drop of more than 30% in
  either aborts the run — nothing is written, committed or pushed. `--force`
  overrides it deliberately.
- **`expand_ranges` rewritten.** Handles `,` and `;` separators, `:` and `-`
  ranges, spaces, reversed ranges, and trailing notes. Refuses to read `3rd Floor
  Doors` as unit 3. Anything it cannot parse produces a warning instead of a
  silent drop.
- **Building matrix located by scanning for floor labels**, not by assuming row 7
  and a 3-row step. Columns are scanned to the end of the sheet and each unit id
  is checked against its floor, so inserted rows/columns no longer break it.
  Duplicate floor blocks are ignored with a warning.
- **Sheet selection is ranked most-specific-first**, so a new sheet called
  "Container Notes" or reordering the tabs cannot hijack the schedule or the
  transit figures.
- **Unrecognised statuses map to `HOLD`** ("On Hold", grey) instead of silently
  becoming "Projected". `CANCELLED`, `CUSTOMS`, `DELAYED`, `ON HOLD`, `TBD` are
  mapped explicitly.
- **Header aliases widened** (`CTNR #`, `CONT #`, `DELIVERY ETA`, …) and a missing
  column now produces a warning naming the field.
- **Dates** emit a machine-readable `deliveryISO` alongside the display string.
- **MISC NOTES** is read into a `note` field (prose only — bare numbers ignored).
- **Warnings** print to the console and to `111W_update_log.txt`.

**`111W_Container_Dashboard.html`**

- Total Containers, "of NNN" kitchens, and the "Floors 3–12" heading are all
  derived from the data. No hard-coded counts remain.
- **Floor progress counts only units that have a kitchen.** With every container
  delivered, floors 3–10 now reach **100%** (verified). Floors with no kitchens
  at all are no longer rendered, so the 17 permanent "0/58" rows are gone.
- **Units Delivered counts kitchens off the matrix**, so unit-level overrides
  (partial deliveries) move the headline number.
- **Next Up has a "Past due" section** with a day count, plus a line for
  containers with no firm date. Late rows are also flagged in the table.
- **Dates parsed as local**, fixing the off-by-one (`2026-08-15` rendered as Aug 14).
- **All spreadsheet values are HTML-escaped** on output, including `title`
  attributes. Matrix tooltips now name the owning container.
- Unit column count and the damage-report dropdown are derived from the data
  rather than fixed at 58 / floors 3–29.
- Empty states for the table and the matrix.

**`run_update.bat`** — passes through arguments, reports a failed run explicitly,
and `pause`s so the window no longer closes on a crash.

## Verified

All 50 scenarios re-run against the fixed build: no crashes, no JS errors, no
`undefined`/`NaN` in output. Specifically confirmed:

| Was | Now |
|---|---|
| Semicolon in Units crashed the script | Parses, 33 containers |
| `CTNR #` header → 0 containers | 33 containers |
| Extra "Container Notes" sheet → 0 containers | 33 containers |
| Row inserted in matrix → 0 floors | 27 floors |
| Column inserted in matrix → 9 units lost | 0 units lost |
| Tabs reordered → garbage transit figures | Correct figures |
| Text date `2026-08-15` → "Aug 14" | "Aug 15" |
| 6 overdue containers → invisible | "Past due · 6", "26d late" |
| Unit override → KPI unchanged | KPI moves |
| Empty spreadsheet → published and pushed | Publish blocked |
| All delivered → floors capped at 91–97% | Floors 3–10 at 100% |

## Still open — these are data issues, not code

The script now names them on every run; they need a change in the master file:

1. **13 kitchens are in no container:** 1150, 1151, 1248–1258. Floors 11 and 12
   cannot reach 100% until these are assigned.
2. **The `Doors` container's Units cell reads "3rd Floor Doors."** That is a label,
   not a unit list, so the container carries no units. Fine if intentional — if
   those doors map to specific units, put the unit numbers in the cell.

