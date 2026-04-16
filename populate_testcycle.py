"""
populate_testcycle.py
─────────────────────
Updates all header/summary columns in testcycle from my_master.xlsx.
The 36 move-sequence columns (Train ID / Start Location / End Location / Dwell Time × 9)
are LEFT UNTOUCHED — they were manually built and cannot be derived automatically.

Columns updated (by position in the 49-column sheet):
  [1]  Current Day Cycle
  [2]  Number of Cars
  [3]  Revenue Runs
  [4]  Total Deadhead Runs
  [5]  Total Runs
  [6]  Starting Outlying Point
  [7]  First Move             ← weekday only; Sa/Su preserved
  [44] Ending Outlying Point
  [45] Final Move             ← weekday only; Sa/Su preserved
  [46] Next day cycle
  [47] CMF Train?
  [48] CMF Arrival Time
  [49] CMF Departure Time

Columns PRESERVED exactly (never touched):
  [8–43]  Train ID / Start Location / End Location / Dwell Time × 9

Usage:
  python populate_testcycle.py
      --master   my_master.xlsx
      --template testcycle.xlsx
      --output   testcycle_updated.xlsx
"""

import argparse, re, math
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── helpers ───────────────────────────────────────────────────────────────────
def na(v):
    """Return empty string for None/NaN, otherwise the value."""
    if v is None: return ""
    try:
        if math.isnan(float(v)): return ""
    except: pass
    return v

def clean(v):
    """Strip and return string, empty string if blank."""
    s = str(na(v)).strip()
    return "" if s.lower() in ("nan","none","nat") else s

def to_int(v):
    """Convert to int if possible, else empty string."""
    s = clean(v)
    if s.replace('.','').isdigit():
        return int(float(s))
    return ""

# ── column positions (1-based) ────────────────────────────────────────────────
COL_CYCLE      = 1
COL_CARS       = 2
COL_REV        = 3
COL_DH         = 4
COL_TOT        = 5
COL_START      = 6
COL_FIRST_MOVE = 7
# cols 8-43 = move sequences (PRESERVED)
COL_END        = 44
COL_FINAL_MOVE = 45
COL_NEXT       = 46
COL_CMF        = 47
COL_CMF_ARR    = 48
COL_CMF_DEP    = 49

HEADER_COLS = {COL_CYCLE, COL_CARS, COL_REV, COL_DH, COL_TOT,
               COL_START, COL_FIRST_MOVE, COL_END, COL_FINAL_MOVE,
               COL_NEXT, COL_CMF, COL_CMF_ARR, COL_CMF_DEP}

# ── build lookup tables from master ──────────────────────────────────────────
def build_lookups(master_path):
    df_cw  = pd.read_excel(master_path, sheet_name='cycles_weekday')
    df_cwk = pd.read_excel(master_path, sheet_name='cycles_weekend')

    # ── weekday ───────────────────────────────────────────────────────────────
    wd = {}
    for _, r in df_cw.iterrows():
        c = na(r['cycle'])
        if c == "": continue
        cnum = int(float(c))

        cmf_arr = clean(r['cmf_arrive_time'])
        cmf_dep = clean(r['cmf_depart_time'])

        # "Turn:XXX" means the cycle DOES visit CMF but via a turn move (no fixed layover time).
        # These cycles are CMF=Yes; store the turn train reference in the time cells.
        is_turn_arr = cmf_arr.lower().startswith("turn")
        is_turn_dep = cmf_dep.lower().startswith("turn")

        if is_turn_arr: cmf_arr = cmf_arr   # keep "Turn:XXX" as the cell value
        if is_turn_dep: cmf_dep = cmf_dep

        # CMF=Yes if there is a real time OR a Turn annotation (cycle visits CMF either way)
        cmf = "Yes" if (cmf_arr or cmf_dep) else "No"
        dh  = 2 if cmf == "Yes" else 0
        rev = to_int(r['weekday_trips'])
        tot = (rev if isinstance(rev, int) else 0) + dh

        first = clean(r['depart_layover']) or clean(r['depart_cmf'])
        final = clean(r['arrive_layover']) or clean(r['arrive_cmf'])
        layover = clean(r['layover_location'])

        wd[cnum] = dict(
            cars=na(r['cars']), rev=rev, dh=dh,
            tot=tot if tot > 0 else "",
            start=layover, end=layover,
            first=first, final=final,
            next=na(r['next_cycle']),
            cmf=cmf, cmf_arr=cmf_arr, cmf_dep=cmf_dep,
        )

    # ── weekend (Sa / Su) ─────────────────────────────────────────────────────
    wknd = {}
    for _, r in df_cwk.iterrows():
        bc      = int(float(na(r['weekday_cycle'])))
        layover = clean(r['layover_location'])
        cars    = na(r['cars'])
        sa_rev  = to_int(r['sa_trip_count'])
        su_rev  = to_int(r['su_trip_count'])

        # Su next-cycle = the weekday cycle's next_cycle
        wk_next = clean(wd.get(bc, {}).get('next', ''))

        wknd[bc] = {
            'Sa': dict(cars=cars, rev=sa_rev, dh=0,
                       tot=sa_rev if isinstance(sa_rev,int) else "",
                       start=layover, end=layover,
                       next=f"{bc}Su",
                       cmf="No", cmf_arr="N/A", cmf_dep="N/A"),
            'Su': dict(cars=cars, rev=su_rev, dh=0,
                       tot=su_rev if isinstance(su_rev,int) else "",
                       start=layover, end=layover,
                       next=wk_next,
                       cmf="No", cmf_arr="N/A", cmf_dep="N/A"),
        }

    return wd, wknd

def get_fields(cid, wd, wknd):
    """
    Return a dict of updated fields for a given cycle ID string.
    Returns None if cycle has no master data (suspended/empty cycles).
    """
    m2 = re.match(r'^(\d+)(Sa|Su)$', str(cid).strip())
    if m2:
        bc  = int(m2.group(1))
        var = m2.group(2)
        rec = wknd.get(bc, {}).get(var)
        # Sa/Su: first_move and final_move come from EXISTING testcycle (not updated)
        if rec:
            rec = dict(rec)
            rec['preserve_times'] = True   # flag: don't update first/final move
        return rec
    else:
        try:
            cnum = int(float(str(cid).strip()))
            rec  = wd.get(cnum)
            if rec: rec = dict(rec); rec['preserve_times'] = False
            return rec
        except:
            return None

# ── styling helpers ───────────────────────────────────────────────────────────
HDR_FILL    = PatternFill("solid", fgColor="1F3864")
HDR_FONT    = Font(name="Arial", bold=True, color="FFFFFF", size=10)
ALT_FILL    = PatternFill("solid", fgColor="EEF2F7")
WHITE_FILL  = PatternFill("solid", fgColor="FFFFFF")
DATA_FONT   = Font(name="Arial", size=10)
THIN        = Side(style="thin", color="CCCCCC")
BORDER      = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER      = Alignment(horizontal="center", vertical="center")
# Sketchy-data highlights
TBD_FILL    = PatternFill("solid", fgColor="FFD966")   # amber       — TBD trip count
TURN_FILL   = PatternFill("solid", fgColor="FFE699")   # light amber — CMF turn move
SUSP_FILL   = PatternFill("solid", fgColor="F4CCCC")   # light red   — suspended cycle
WARN_FONT   = Font(name="Arial", size=10, bold=True, color="7F4F00")

def style_header_row(ws, n_cols):
    for c in range(1, n_cols+1):
        cell = ws.cell(row=1, column=c)
        cell.fill = HDR_FILL; cell.font = HDR_FONT
        cell.border = BORDER; cell.alignment = CENTER
    ws.row_dimensions[1].height = 30

def style_data_row(ws, row_num, n_cols):
    fill = ALT_FILL if row_num % 2 == 0 else WHITE_FILL
    for c in range(1, n_cols+1):
        cell = ws.cell(row=row_num, column=c)
        cell.fill = fill; cell.font = DATA_FONT
        cell.border = BORDER; cell.alignment = CENTER
    ws.row_dimensions[row_num].height = 18

# ── main update logic ─────────────────────────────────────────────────────────
def populate(master_path, template_path, output_path):
    print("Loading master ...")
    wd, wknd = build_lookups(master_path)
    print(f"  Weekday cycles  : {len(wd)}")
    print(f"  Weekend cycle groups: {len(wknd)}")

    print(f"Opening template: {template_path}")
    wb = load_workbook(template_path)
    ws = wb["Weekday"]
    n_cols = ws.max_column  # should be 49

    print(f"  Sheet columns   : {n_cols}")
    print(f"  Sheet data rows : {ws.max_row - 1}")

    # Re-apply header styling (preserves values, refreshes formatting)
    style_header_row(ws, n_cols)

    updated = 0; skipped = 0; preserved = 0

    for row_num in range(2, ws.max_row + 1):
        cid_cell = ws.cell(row=row_num, column=COL_CYCLE)
        cid = cid_cell.value
        if cid is None: continue

        fields = get_fields(cid, wd, wknd)

        if fields is None:
            # Suspended / empty cycle — highlight entire row in light red
            style_data_row(ws, row_num, n_cols)
            for c in range(1, n_cols+1):
                cell = ws.cell(row=row_num, column=c)
                cell.fill = SUSP_FILL
            # Add a note to the cycle cell so the reason is visible
            ws.cell(row=row_num, column=COL_CYCLE).font = WARN_FONT
            skipped += 1
            continue

        preserve_times = fields.get('preserve_times', False)

        # Update each header column
        updates = {
            COL_CARS:       fields['cars'],
            COL_REV:        fields['rev'],
            COL_DH:         fields['dh'],
            COL_TOT:        fields['tot'],
            COL_START:      fields['start'],
            COL_END:        fields['end'],
            COL_NEXT:       fields['next'],
            COL_CMF:        fields['cmf'],
            COL_CMF_ARR:    fields['cmf_arr'],
            COL_CMF_DEP:    fields['cmf_dep'],
        }
        if not preserve_times:
            updates[COL_FIRST_MOVE] = fields['first']
            updates[COL_FINAL_MOVE] = fields['final']
        else:
            preserved += 2  # Sa/Su first/final kept as-is

        for col, val in updates.items():
            ws.cell(row=row_num, column=col).value = val if val != "" else None

        style_data_row(ws, row_num, n_cols)

        # ── Per-cell sketchy highlights (applied AFTER base row styling) ────
        # TBD trips: amber highlight on Revenue Runs and Total Runs
        if fields['rev'] == "TBD" or str(fields.get('rev','')).upper() == "TBD":
            for c in [COL_REV, COL_DH, COL_TOT]:
                cell = ws.cell(row=row_num, column=c)
                cell.fill = TBD_FILL; cell.font = WARN_FONT

        # Turn-move CMF: amber on CMF columns to flag that time is not a real clock time
        cmf_arr_val = str(fields.get('cmf_arr','') or '')
        cmf_dep_val = str(fields.get('cmf_dep','') or '')
        if cmf_arr_val.startswith('Turn') or cmf_dep_val.startswith('Turn'):
            for c in [COL_CMF, COL_CMF_ARR, COL_CMF_DEP]:
                cell = ws.cell(row=row_num, column=c)
                cell.fill = TURN_FILL; cell.font = WARN_FONT

        updated += 1

    # Set column widths
    col_widths = [14,10,11,16,10,18,12] + [13,14,12,14]*9 + [18,12,16,10,16,18]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "H2"  # freeze up to move columns

    # ── Weekend sheet ─────────────────────────────────────────────────────────
    print("Populating Weekend sheet ...")
    wknd_count = populate_weekend(wb, wd, wknd)

    wb.save(output_path)

    print(f"\nDone → {output_path}")
    print(f"  Rows updated     : {updated}")
    print(f"  Rows skipped     : {skipped}  (suspended/empty cycles)")
    print(f"  Sa/Su times kept : {preserved} cells preserved from existing file")
    print(f"  Move sequences   : ALL preserved (cols 8–43 untouched)")

# ── validation report ─────────────────────────────────────────────────────────
def validate(output_path, master_path):
    print("\n── Data check ──────────────────────────────────────────────")
    xf = pd.ExcelFile(output_path)
    print(f"  Sheets           : {xf.sheet_names}")

    df = pd.read_excel(output_path, sheet_name="Weekday")
    print(f"\n  Weekday sheet:")
    print(f"    Total rows     : {len(df)}")
    print(f"    Rows with cars : {df['Number of Cars'].notna().sum()}")
    print(f"    CMF = Yes      : {(df['CMF Train?']=='Yes').sum()}")
    print(f"    CMF = No       : {(df['CMF Train?']=='No').sum()}")
    has_moves = df['Train ID'].notna()
    print(f"    Rows with moves: {has_moves.sum()}")
    no_moves = df[~has_moves]['Current Day Cycle'].tolist()
    if no_moves: print(f"    No moves (expected): {no_moves}")

    if "Weekend" in xf.sheet_names:
        dw = pd.read_excel(output_path, sheet_name="Weekend")
        print(f"\n  Weekend sheet:")
        print(f"    Total rows     : {len(dw)}")
        print(f"    Rows with cars : {dw['Number of Cars'].notna().sum()}")
        has_moves_w = dw['Train ID'].notna()
        print(f"    Rows with moves: {has_moves_w.sum()}")
        print(f"    Cycle IDs      : {dw['Current Day Cycle'].tolist()}")

    print("────────────────────────────────────────────────────────────")

# ── entry point ───────────────────────────────────────────────────────────────
def main():
    p = argparse.ArgumentParser()
    p.add_argument("--master",   required=True)
    p.add_argument("--template", required=True)
    p.add_argument("--output",   required=True)
    args = p.parse_args()
    populate(args.master, args.template, args.output)
    validate(args.output, args.master)

# ═══════════════════════════════════════════════════════════════════════════════
# WEEKEND SHEET — populate Sheet1 with Sa/Su rows (same 49-col format)
# ═══════════════════════════════════════════════════════════════════════════════

def populate_weekend(wb, wd, wknd):
    """
    Renames Sheet1 to 'Weekend' and populates it with all Sa/Su rows
    from the Weekday sheet, updating header/summary columns from master
    while preserving move sequences exactly.
    """
    ws_wd  = wb["Weekday"]
    ws_wknd = wb["Sheet1"]
    ws_wknd.title = "Weekend"

    n_cols = ws_wd.max_column  # 49

    # ── copy header row from Weekday ──────────────────────────────────────────
    for c in range(1, n_cols + 1):
        src  = ws_wd.cell(row=1, column=c)
        dest = ws_wknd.cell(row=1, column=c)
        dest.value     = src.value
        dest.fill      = HDR_FILL
        dest.font      = HDR_FONT
        dest.border    = BORDER
        dest.alignment = CENTER
    ws_wknd.row_dimensions[1].height = 30

    # ── collect Sa/Su rows from Weekday ───────────────────────────────────────
    sa_su_rows = []
    for row_num in range(2, ws_wd.max_row + 1):
        cid = ws_wd.cell(row=row_num, column=COL_CYCLE).value
        if cid is None:
            continue
        if re.search(r'(Sa|Su)$', str(cid).strip()):
            # Read entire row
            row_data = [ws_wd.cell(row=row_num, column=c).value for c in range(1, n_cols + 1)]
            sa_su_rows.append(row_data)

    # ── write to weekend sheet ────────────────────────────────────────────────
    for dest_row, row_data in enumerate(sa_su_rows, 2):
        cid = str(row_data[COL_CYCLE - 1]).strip()
        fields = get_fields(cid, wd, wknd)

        # Write all values first (preserves moves)
        for c, val in enumerate(row_data, 1):
            ws_wknd.cell(row=dest_row, column=c).value = val

        # Then overlay the header/summary columns from master
        if fields:
            updates = {
                COL_CARS:    fields['cars'],
                COL_REV:     fields['rev'],
                COL_DH:      fields['dh'],
                COL_TOT:     fields['tot'],
                COL_START:   fields['start'],
                COL_END:     fields['end'],
                COL_NEXT:    fields['next'],
                COL_CMF:     fields['cmf'],
                COL_CMF_ARR: fields['cmf_arr'],
                COL_CMF_DEP: fields['cmf_dep'],
                # Sa/Su: preserve First Move and Final Move from existing data
            }
            for col, val in updates.items():
                ws_wknd.cell(row=dest_row, column=col).value = val if val != "" else None

        style_data_row(ws_wknd, dest_row, n_cols)

        # ── Per-cell sketchy highlights ───────────────────────────────────
        if fields:
            if fields.get('rev') == "TBD" or str(fields.get('rev','')).upper() == "TBD":
                for c in [COL_REV, COL_DH, COL_TOT]:
                    cell = ws_wknd.cell(row=dest_row, column=c)
                    cell.fill = TBD_FILL; cell.font = WARN_FONT

    # ── column widths (same as Weekday) ───────────────────────────────────────
    col_widths = [14,10,11,16,10,18,12] + [13,14,12,14]*9 + [18,12,16,10,16,18]
    for i, w in enumerate(col_widths, 1):
        ws_wknd.column_dimensions[get_column_letter(i)].width = w

    ws_wknd.freeze_panes = "H2"

    print(f"  Weekend sheet: {len(sa_su_rows)} rows written")
    return len(sa_su_rows)

if __name__ == "__main__":
    main()

