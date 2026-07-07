#!/usr/bin/env python3
"""Schedule-deck generator (generic, block-based).

Model — nothing here is tied to any sample-specific names:

  * PRODUCT master row (제품 일정).
  * One or more EVALUATION BLOCKS, each with a title. Every block is read the
    same way: a title plus one or more EVALUATION ITEMS (평가 항목). Each item
    owns one table row and one lane of milestone markers.

All blocks live in identical 5-column groups starting at column C and stepping by
5 (C-G, H-L, M-Q, ...), so blocks/items/lanes are sized purely from the data: any
number of blocks, any number of items per block, any number of check-posts per
item. Markers sit at their true date x; the timeline is fitted to the data, drops
empty months and widens busy ones.
"""
import copy
import os
import sys
import calendar
import datetime as dt
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import openpyxl
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
import engine as E

UNIT = E.UNIT
TABLE_TOP = 134.0
HDR_H = 28.8
ITEM_ROW_H = 59.1                 # uniform evaluation-item row height
TABLE_BOTTOM = 712.0
GAP_BELOW_SUBTITLE = 8.0          # px: final table top sits this far below the 부제 box bottom
FIRST_BLOCK_COL0 = 2              # first evaluation block starts at column C (0-based 2)
BLOCK_W = 5                      # group | target-label | target-date | check | check-date
MAX_BLOCKS = 21
PRODUCT_SLOTS = 5                # template's pre-built product markers (auto-extended)
MONTH_COLS = 8                   # physical month columns in the template table

HL_RGB = {"orange": RGBColor(0xFF, 0xA5, 0x00), "skyblue": RGBColor(0x87, 0xCE, 0xEB)}
COL0_W = 54.0
COL1_W = 44.0
STATUS_W = 58.0
LABEL_PT = 10.0
STATUS_PT = 9.0
STATUS_STYLE = {"완료": ("C6EFCE", "006100"), "진행중": ("FFEB9C", "9C6500"),
                "예정": ("DDEBF7", "1F4E79")}
CHAR_PT = 5.3
CELL_MARGIN_PT = 1.0
SUBCOL_FONT_FLOOR = 8.0          # smallest 평가항목명 font before widening the column
SUBCOL_FONT_BASE = 12.0         # calibration base for CHAR_PT
COL1_FONT_MAX = 10.0            # max render font for the item-name column (keeps current visible size)
SUBCOL_FIT_SAFETY = 0.90        # headroom vs PowerPoint real font metrics (text counts as fitting only within this fraction)
COL1_W_MAX = 84.0               # 평가항목명 열 최대 폭(극단적으로 긴 이름에서만 확대)
LABEL_MAX_W = 140.0             # 체크포스트 라벨 박스 최대 폭(pt) — 그 이상은 폰트 축소
LABEL_FONT = 10.0
LABEL_FONT_FLOOR = 8.0
# Label-aware timeline layout: space markers so a label centred on its circle never overlaps
# the next, widen busy months to make room, and shrink the label font only if the whole
# timeline still cannot hold them.
GAP_LABEL = 6.0                 # clear gap (pt) required between two adjacent label glyph boxes
CIRCLE_MIN_SEP = 34.0           # min centre-to-centre so circles (~30pt) never touch
EDGE_PAD = 12.0                 # padding inside the timeline frame at both ends
GLYPH_SAFETY = 1.12             # inflate the estimated glyph width a touch (real metrics vary)
BOX_SAFETY = 1.30               # label box is wider than the glyphs so it never wraps
LABEL_FS_FLOOR = 0.50           # smallest label-font scale before we stop shrinking
EDGE_PAD_M = 8.0                # padding inside each month column
MONTH_MIN_W = 46.0              # a month column never narrower than this
PRODUCT_LAST_RGB = RGBColor(0x00, 0x64, 0x00)   # 제품 일정 마지막 마일스톤 = 짙은 녹색

# Generic marker source shapes cloned for every lane (the template happens to name
# them with Korean strings; the names carry no special meaning).
GEN_CIRCLE = E.TEMPLATE_NAMES["gen_circle"]   # 타원 33
GEN_LABEL = E.TEMPLATE_NAMES["gen_label"]     # 직사각형 32

# original geometry, snapshotted so install_geometry() is idempotent across runs
_ORIG_GEOM = None
_ORIG_ANCHORS = None


# --------------------------------------------------------------------------- #
# Workbook reading
# --------------------------------------------------------------------------- #
def read_values_ext(ws, nrows=16, ncols=110):
    vals = []
    for r in range(1, nrows + 1):
        row = []
        for c in range(1, ncols + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, dt.datetime):
                v = E.date_to_excel_serial(v.date())
            elif isinstance(v, dt.date):
                v = E.date_to_excel_serial(v)
            row.append(v)
        vals.append(row)
    return vals


def group_target_dates(V, rows, group_col, tlabel_col, tdate_col):
    """Map each evaluation item (group) -> its OWN 목표일정 dates, tracking the
    current group label down the block so a neighbour's target never bleeds its
    colour onto this item."""
    out, cur = {}, None
    for r in rows:
        gl = E.readable_cell_value(E._cell(V, r, group_col))
        if gl:
            cur = gl
            out.setdefault(cur, set())
        tl = E.readable_cell_value(E._cell(V, r, tlabel_col))
        td = E._cell(V, r, tdate_col)
        if cur is not None and tl and E.is_number(td):
            out.setdefault(cur, set()).add(td)
    return out


def assign_hl(events, target_dates):
    """Per item: the LAST check-post is skyblue (완료); a check-post matching this
    item's own 목표일정 is orange; when 목표 == 마지막 그 마일스톤은 하늘색
    (skyblue wins). Everything else stays uncoloured."""
    dates = [e["date"] for e in events if E.is_number(e["date"])]
    mx = max(dates) if dates else None
    for e in events:
        if E.is_number(e["date"]) and e["date"] == mx:
            e["hl"] = "skyblue"
        elif E.is_number(e["date"]) and e["date"] in target_dates:
            e["hl"] = "orange"
        else:
            e["hl"] = None


def compute_status(events, today):
    dates = [e["date"] for e in events if E.is_number(e["date"])]
    if not dates:
        return None
    passed = [d for d in dates if d <= today]
    if len(passed) == len(dates):
        return "완료"
    return "진행중" if passed else "예정"


def _read_grid_openpyxl(path, sheet=None):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet] if (sheet and sheet in wb.sheetnames) else wb.worksheets[0]
    nrows = min(max(16, ws.max_row), 500)
    return read_values_ext(ws, nrows=nrows), nrows


def _com_cell(v):
    """Normalise an Excel-COM cell value to match read_values_ext (dates -> serial)."""
    if isinstance(v, dt.datetime):
        return E.date_to_excel_serial(v.date())
    if isinstance(v, dt.date):
        return E.date_to_excel_serial(v)
    return v


def _read_grid_com(path, sheet=None, ncols=110, maxrows=500):
    """사내 문서보안(DRM)/CDR 폴백: openpyxl 이 raw 파일을 zip 으로 못 열 때 Excel(COM)로 값을 읽는다.
    comtypes(순수 파이썬 COM · late-binding)로 구현해 PyInstaller 번들이 안정적이다. (Excel 설치 필요)
    엔진 직렬값 기준(EPOCH=1899-12-30)이 Excel 기본 직렬값과 같아 그대로 사용 가능하다."""
    import comtypes.client
    excel = comtypes.client.CreateObject("Excel.Application", dynamic=True)
    try:
        for prop, val in (("Visible", False), ("DisplayAlerts", False), ("AutomationSecurity", 3)):
            try:
                setattr(excel, prop, val)
            except Exception:
                pass
        wb = excel.Workbooks.Open(os.path.abspath(path), 0, True)   # (Filename, UpdateLinks, ReadOnly)
        try:
            ws = wb.Worksheets(sheet) if sheet else wb.Worksheets(1)
            used = ws.UsedRange
            last_row = int(used.Row) + int(used.Rows.Count) - 1
            nrows = min(max(16, last_row), maxrows)
            val = ws.Range(ws.Cells(1, 1), ws.Cells(nrows, ncols)).Value
            data = val() if callable(val) else val      # comtypes dynamic: .Value 는 호출형(NamedProperty)
            V = []
            for r in range(nrows):
                rowvals = data[r] if (data and r < len(data)) else None
                row = []
                for c in range(ncols):
                    v = rowvals[c] if (rowvals is not None and c < len(rowvals)) else None
                    row.append(_com_cell(v))
                V.append(row)
            return V, nrows
        finally:
            try:
                wb.Close(False)
            except Exception:
                pass
    finally:
        try:
            excel.Quit()
        except Exception:
            pass


def parse_input(path, sheet=None):
    try:
        V, nrows = _read_grid_openpyxl(path, sheet)
    except Exception as e_open:
        try:
            V, nrows = _read_grid_com(path, sheet=sheet)      # 사내 DRM/CDR 파일은 Excel 경유로 재시도
        except Exception as e_com:
            raise RuntimeError(
                "입력 파일을 열 수 없습니다. 사내 문서보안(DRM)으로 보호된 파일일 수 있습니다. "
                "Excel 경유 읽기도 실패했습니다(Excel 미설치 또는 보안 정책 차단). "
                f"[openpyxl: {e_open}] [Excel: {e_com}]"
            )
    data_rows = list(range(6, nrows))            # 0-based: sheet row 7+
    prod_rows = list(range(5, nrows))            # 0-based: sheet row 6+
    product = E.read_items(V, prod_rows, 0, 1)

    blocks = []
    for b in range(MAX_BLOCKS):
        col = FIRST_BLOCK_COL0 + b * BLOCK_W
        lane = E.read_subcolumn_lane(V, data_rows, col, col + 1, col + 2, col + 3, col + 4)
        targets = group_target_dates(V, data_rows, col, col + 1, col + 2)
        items = []
        for g in lane["groups"]:               # one item per group-label row (same-named items stay separate)
            if not g["name"]:                  # no label -> let the 'Plan title only' merged fallback handle it
                continue
            evs = [e for e in g["events"] if e["source"] == "check"]
            assign_hl(evs, targets.get(g["name"], set()))
            items.append({"name": g["name"], "events": evs})
        title = E._cell(V, 3, col) or E._cell(V, 4, col)   # block title: 4행(신양식) 우선, 5행(구양식) 폴백
        title_s = str(title).strip() if title else ""
        if not items and title_s:
            # 평가항목 세분화 없이 상위 Plan만 있는 경우: 제목을 단일 항목으로 -> col0+col1 병합 행으로 표시
            evs = [e for e in lane["events"] if e["source"] == "check"]
            assign_hl(evs, targets.get("", set()))
            items.append({"name": title_s, "events": evs, "merged": True})
        if items:
            blocks.append({"title": title_s, "items": items})

    return {"title": E._cell(V, 1, 1) or "", "punchline": E._cell(V, 2, 1) or "",
            "product_title": E._cell(V, 3, 0) or "",   # A4: 제품행 라벨(사용자 편집)
            "product": product, "blocks": blocks}


# --------------------------------------------------------------------------- #
# Timeline: always fitted to the data, empty months dropped, busy months widened
# --------------------------------------------------------------------------- #
def build_scale(all_dates):
    """Fitted scale + the list of month-keys to actually show (non-empty months,
    chronological, capped at MONTH_COLS) with a weight per month for width."""
    mn, mx = min(all_dates), max(all_dates)
    start = E.month_start_serial(mn)
    end = E.month_end_serial(mx)
    keys = E.month_keys_between(start, end)
    counts = {k: 0 for k in keys}
    for d in all_dates:
        k = E.month_key(d)
        if k in counts:
            counts[k] += 1
    active = [k for k in keys if counts[k] > 0]          # every month that has a marker
    weights = [counts[k] + 1.0 for k in active]          # busy months get more width
    scale = {"type": "fitted", "start_serial": start, "end_serial": end,
             "header_month_keys": active}
    return scale, weights


def install_geometry(weights, col1_w=COL1_W):
    """Shift the timeline right for the 현황 column and hand each shown month a
    width proportional to its marker count (empty months collapse to 0)."""
    global _ORIG_GEOM, _ORIG_ANCHORS
    if _ORIG_GEOM is None:
        _ORIG_GEOM = dict(E.TABLE_GEOMETRY)
        _ORIG_ANCHORS = [list(a) for a in E.TIMELINE_ANCHORS]
    g = _ORIG_GEOM
    left = g["left"]
    W = list(g["column_widths"])
    old_tl_left = left + W[0] + W[1]
    table_right = left + sum(W)
    new_tl_left = left + COL0_W + col1_w + STATUS_W
    frame = table_right - new_tl_left
    tx = lambda x: new_tl_left + (x - old_tl_left) * (frame / (table_right - old_tl_left))
    E.TIMELINE_ANCHORS = [[s, tx(x)] for s, x in _ORIG_ANCHORS]
    wsum = sum(weights) or 1.0
    month_widths = [frame * w / wsum for w in weights]
    month_widths += [0.0] * max(0, MONTH_COLS - len(month_widths))  # keep >=8 for the template grid
    E.TABLE_GEOMETRY = dict(g, column_widths=[COL0_W, col1_w, STATUS_W] + month_widths,
                            month_start_column=3, month_column_count=max(1, len(weights)))
    return tx


def _day_frac(serial):
    """Position of a date inside its own month: day 1 -> 0.0, last day -> 1.0."""
    d = E.excel_serial_to_date(serial)
    dim = calendar.monthrange(d.year, d.month)[1]
    return (d.day - 1) / max(1, dim - 1)


def compute_timeline_layout(product, blocks, active_months, col1_w):
    """Place every unique date so that (1) markers stay in date order and inside their own
    month column, (2) adjacent markers are far enough apart that a label centred on its circle
    never overlaps the next, and (3) busy / wide-label months get proportionally more width.
    If the whole timeline still cannot hold the labels, the label font is scaled down until it
    fits. Returns {centers: {serial: x}, month_widths: [..aligned to active_months..],
    font_scale: fs}."""
    g = _ORIG_GEOM if _ORIG_GEOM is not None else E.TABLE_GEOMETRY
    left = g["left"]
    widths = list(g["column_widths"])
    frame_left = left + COL0_W + col1_w + STATUS_W
    frame_right = left + sum(widths)
    frame = frame_right - frame_left

    by_date = {}
    for it in product:
        if E.is_number(it["date"]):
            by_date.setdefault(it["date"], []).append(str(it["label"]))
    for b in blocks:
        for itm in b["items"]:
            for e in itm["events"]:
                if E.is_number(e["date"]):
                    by_date.setdefault(e["date"], []).append(str(e["label"]))
    dates = sorted(by_date)
    if not dates:
        return {"centers": {}, "month_widths": [frame], "font_scale": 1.0}

    by_month = {}
    for d in dates:
        by_month.setdefault(E.month_key(d), []).append(d)
    months = [m for m in active_months if m in by_month] or sorted(by_month)

    def _seps(ds, Wd):
        return [max((Wd[ds[i]] + Wd[ds[i + 1]]) / 2.0 + GAP_LABEL, CIRCLE_MIN_SEP)
                for i in range(len(ds) - 1)]

    # shrink the label font only as far as needed for every MONTH to hold its own markers
    fs = 1.0
    Wd = None
    month_need = []
    for _ in range(24):
        Wd = {d: max((_label_glyph_width(t, fs) for t in by_date[d]), default=10.0) for d in dates}
        month_need = []
        for m in months:
            ds = by_month[m]
            if len(ds) == 1:
                need_m = Wd[ds[0]] + 2 * EDGE_PAD_M
            else:
                need_m = Wd[ds[0]] / 2.0 + sum(_seps(ds, Wd)) + Wd[ds[-1]] / 2.0 + 2 * EDGE_PAD_M
            month_need.append(max(need_m, MONTH_MIN_W))
        if sum(month_need) <= frame or fs <= LABEL_FS_FLOOR:
            break
        fs = max(LABEL_FS_FLOOR, fs * frame / sum(month_need) * 0.995)

    # month-column widths: each month's need scaled so the whole row fills the frame
    # (busy / wide-label months get proportionally more width).
    scale_w = frame / max(1e-6, sum(month_need))
    month_widths = [mn * scale_w for mn in month_need]
    edges = [frame_left]
    for w in month_widths:
        edges.append(edges[-1] + w)

    # Place markers WITHIN their own month column at the true day-of-month (1st near the front, last
    # day near the back), then push apart only where two would overlap. Keeping each month's markers
    # inside its column is essential: a crowded month is compressed within ITS column and never shifts
    # another month's markers (so the master/product lane keeps every marker in the right month).
    centers = {}
    for mi, m in enumerate(months):
        cl, cr = edges[mi], edges[mi + 1]
        ds = by_month[m]
        n = len(ds)
        lo_b = cl + EDGE_PAD_M                               # clamp the CIRCLE (not the label box)
        hi_b = cr - EDGE_PAD_M                               # so markers reach their true day position
        span = max(1e-6, hi_b - lo_b)
        des = [lo_b + _day_frac(d) * span for d in ds]      # day-of-month proportional, full column
        if n == 1:
            x = [des[0]]                                    # single marker -> at its true day position
        else:
            # Pick the spacing that lets markers keep their TRUE day position. First try full label
            # spacing (labels never touch). If that needs more than the column, fall back to CIRCLE
            # spacing (circles never touch, labels may) -- this keeps the 20th/25th/28th near the back
            # instead of collapsing to equal steps. Only if even circles won't fit, space evenly.
            label_seps = _seps(ds, Wd)
            circ_seps = [CIRCLE_MIN_SEP] * (n - 1)
            seps = (label_seps if sum(label_seps) <= span
                    else circ_seps if sum(circ_seps) <= span else None)
            if seps is None:                                # genuinely too many even for circles
                step = max(span / (n - 1), CIRCLE_MIN_SEP)
                total = step * (n - 1)
                start = (cl + cr) / 2.0 - total / 2.0
                x = [start + i * step for i in range(n)]
            else:
                # minimal-displacement isotonic placement (PAV) around the day positions: each marker
                # moves the least amount needed to clear the next, never collapsing to equal steps.
                pref = [0.0] * n
                for i in range(1, n):
                    pref[i] = pref[i - 1] + seps[i - 1]
                y = [des[i] - pref[i] for i in range(n)]
                vals, cnt = [], []                          # pool-adjacent-violators (isotonic regression)
                for yi in y:
                    vals.append(yi); cnt.append(1)
                    while len(vals) > 1 and vals[-2] > vals[-1]:
                        v2, k2 = vals.pop(), cnt.pop()
                        v1, k1 = vals.pop(), cnt.pop()
                        vals.append((v1 * k1 + v2 * k2) / (k1 + k2)); cnt.append(k1 + k2)
                gg = []
                for v, k in zip(vals, cnt):
                    gg.extend([v] * k)
                x = [gg[i] + pref[i] for i in range(n)]
                if x[-1] > hi_b:                            # keep an overflowing tail in-column
                    x[-1] = hi_b
                    for i in range(n - 2, -1, -1):
                        x[i] = min(x[i], x[i + 1] - seps[i])
                if x[0] < lo_b:                             # then an overflowing head
                    x[0] = lo_b
                    for i in range(1, n):
                        x[i] = max(x[i], x[i - 1] + seps[i - 1])
        # Hard invariant: every circle stays inside its OWN month column [lo_b, hi_b],
        # so markers never reverse across a month boundary (e.g. 8/2 can never land left of
        # 7/31) regardless of label width. Do it WITHOUT collapsing the spacing: shift the
        # whole cluster in as a rigid body, and only compress (uniformly, order-preserving)
        # if it is genuinely wider than the column -- so intra-month circle separation is kept.
        if x[0] < lo_b:
            x = [xi + (lo_b - x[0]) for xi in x]
        if x[-1] > hi_b:
            x = [xi - (x[-1] - hi_b) for xi in x]
            if x[0] < lo_b:                              # wider than the column: compress to fit
                w = x[-1] - x[0]
                x = ([lo_b + (xi - x[0]) / w * (hi_b - lo_b) for xi in x] if w > 1e-6
                     else [lo_b for _ in x])
        for i, d in enumerate(ds):
            centers[d] = x[i]
    return {"centers": centers, "month_widths": month_widths, "font_scale": fs}


def insert_status_structure(table):
    """Insert the 현황 column (between label cols and the months) into every row,
    then push the geometry's column widths onto the physical grid."""
    tbl = table._tbl
    grid = tbl.find(qn("a:tblGrid"))
    gcs = grid.findall(qn("a:gridCol"))
    grid.insert(list(grid).index(gcs[2]), copy.deepcopy(gcs[2]))
    for ti, tr in enumerate(tbl.findall(qn("a:tr"))):
        tcs = tr.findall(qn("a:tc"))
        ref = tcs[2]
        ntc = copy.deepcopy(ref)
        for a in ("gridSpan", "rowSpan", "hMerge", "vMerge"):
            if ntc.get(a) is not None:
                del ntc.attrib[a]
        for t in ntc.iter(qn("a:t")):
            t.text = ""
        if ti == 1:                      # product master row: no 현황 -> merge into it
            tcs[0].set("gridSpan", "3")
            ntc.set("hMerge", "1")
        tr.insert(list(tr).index(ref), ntc)
    for gc, w in zip(grid.findall(qn("a:gridCol")), E.TABLE_GEOMETRY["column_widths"]):
        gc.set("w", str(int(round(w * UNIT))))


# --------------------------------------------------------------------------- #
# Text helpers (cell wrap, fonts, status)
# --------------------------------------------------------------------------- #
def _disp_w(s):
    return sum(2 if ord(ch) >= 0x1100 else 1 for ch in s)


def wrap_balanced(text, usable_w_units):
    """Wrap each line to at most 2 lines (one break) only when it overflows
    usable_w_units, breaking at a separator (space/'/'/'-') first so a word is
    never split mid-token, then minimising the longest line."""
    seps = " /-·"
    out = []
    for raw in str(text).split("\n"):
        line = raw.strip()
        if not line:
            continue
        if len(line) <= 1 or _disp_w(line) <= usable_w_units:
            out.append(line)
            continue
        best = None
        for i in range(1, len(line)):
            l, r = line[:i].strip(), line[i:].strip()
            if not l or not r:
                continue
            clean = 0 if (line[i - 1] in seps or line[i] in seps) else 1
            key = (clean, max(_disp_w(l), _disp_w(r)), abs(_disp_w(l) - _disp_w(r)))
            if best is None or key < best[0]:
                best = (key, [l, r])
        out.extend(best[1] if best else [line])
    return "\n".join(out)


def set_label_font(label, pt):
    for para in label.sp.text_frame.paragraphs:
        for run in para.runs:
            run.font.size = Pt(pt)


def _label_width_for(text, font_pt):
    """Box width (pt) that holds text on ONE line at font_pt."""
    return _disp_w(str(text)) * (CHAR_PT * font_pt / SUBCOL_FONT_BASE) / 0.75 * 1.4 + 2 * CELL_MARGIN_PT


def _fit_label(text):
    """Pick (font, box_width) so a check-post label stays on one line: keep the
    default font until the box would exceed LABEL_MAX_W, then shrink to a floor."""
    f = LABEL_FONT
    while f > LABEL_FONT_FLOOR and _label_width_for(text, f) > LABEL_MAX_W:
        f -= 0.5
    return f, min(_label_width_for(text, f), LABEL_MAX_W)


def _label_glyph_width(text, fs=1.0):
    """Estimated one-line glyph width of a marker label at the (optionally scaled) font."""
    f, _ = _fit_label(str(text))
    return _label_width_for(str(text), max(LABEL_FONT_FLOOR, f * fs)) * GLYPH_SAFETY


def _center_label_on_circle(lbl, cir, text, fs=1.0):
    """Centre a marker label EXACTLY under its circle (dx == 0). The box is sized to hold the
    text on one line (wrap-safe) and kept symmetric about the circle centre, so it never drifts;
    non-overlap is guaranteed upstream by the date spacing, not by moving labels."""
    text = str(text).replace("\r\n", "\n").replace("\r", "\n")   # CR -> LF: no phantom blank line
    lines = text.split("\n") if "\n" in text else [text]   # honour Alt+Enter line breaks
    while len(lines) > 1 and lines[0].strip() == "":
        lines.pop(0)
    while len(lines) > 1 and lines[-1].strip() == "":
        lines.pop()
    tf = lbl.sp.text_frame
    tf.word_wrap = False
    # CRITICAL: turn OFF auto-fit. With spAutoFit, PowerPoint re-sizes the box to the text on
    # open (anchored left), which shifts the box centre and breaks the circle alignment even
    # though the saved geometry is centred. noAutofit keeps the box exactly where we put it.
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = 0          # no internal L/R padding
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    widest = max(lines, key=lambda s: _disp_w(s))
    f, _ = _fit_label(widest)                       # font that fits the widest line on one line
    f = max(LABEL_FONT_FLOOR, f * fs)
    # rebuild paragraphs: one per explicit line, centred, keeping the run's colour
    first = tf.paragraphs[0]
    for extra in list(tf.paragraphs[1:]):
        extra._p.getparent().remove(extra._p)
    runs = first.runs
    if runs:
        base = runs[0]
        for rr in runs[1:]:
            rr._r.getparent().remove(rr._r)
    else:
        base = first.add_run()
    base.text = lines[0]
    first.alignment = PP_ALIGN.CENTER
    for ln in lines[1:]:
        para = tf.add_paragraph()
        para.alignment = PP_ALIGN.CENTER
        nr = para.add_run()
        nr.text = ln
        E._copy_run_font(base, nr)
    set_label_font(lbl, f)
    box = max((_label_width_for(ln, f) for ln in lines), default=10.0)   # tight box, no L/R padding (margins already 0)
    cc = cir.left + cir.width / 2.0
    lbl.width = box
    lbl.left = cc - box / 2.0
    lbl.height = max(12.0, len(lines) * f * 1.6)    # box height tracks line count & font


def _line_cap_units(width_pt, font_pt):
    """How many display-units (CJK=2, else=1) fit on one line of a width_pt cell
    at font_pt, calibrated so font==SUBCOL_FONT_BASE matches the legacy CHAR_PT."""
    return max(1.0, (width_pt * 0.75 - 2 * CELL_MARGIN_PT) / (CHAR_PT * font_pt / SUBCOL_FONT_BASE))


def _fits_one_line(disp_units, width_pt, font_pt):
    return disp_units <= _line_cap_units(width_pt, font_pt) * SUBCOL_FIT_SAFETY


def _fits_two_lines(disp_units, width_pt, font_pt):
    return disp_units <= 2 * _line_cap_units(width_pt, font_pt) * SUBCOL_FIT_SAFETY


def _name_lines(name):
    """Item-name lines as the user typed them (explicit Alt+Enter breaks)."""
    return str(name).split("\n") or [""]


def compute_col1_width(names):
    """Item-name column width: keep COL1_W normally; widen only up to COL1_W_MAX so names fit
    without forcing an EXTRA line. A name the user PRE-BROKE keeps its own lines, so each line
    must fit on one line; an unbroken name only needs to fit within 2 auto-wrapped lines."""
    w = COL1_W
    for n in names:
        lines = _name_lines(n)
        if len(lines) > 1:                          # user pre-broke -> each line on one line
            need = max((_disp_w(ln) for ln in lines), default=0)
            while w < COL1_W_MAX and not _fits_one_line(need, w, COL1_FONT_MAX):
                w += 2.0
        else:                                       # single string -> within 2 auto-wrapped lines
            need = _disp_w(lines[0])
            while w < COL1_W_MAX and not _fits_two_lines(need, w, COL1_FONT_MAX):
                w += 2.0
    return min(w, COL1_W_MAX)


def _fit_subcolumn_font(text, width_pt):
    """Pick the item-name font. If the user PRE-BROKE the name, fit the WIDEST line on one line
    (never add a break). Otherwise: largest font that fits one line, else <=2 auto-wrapped lines."""
    lines = _name_lines(text)
    if len(lines) > 1:                              # respect explicit breaks
        need = max((_disp_w(ln) for ln in lines), default=0)
        f = COL1_FONT_MAX
        while f > SUBCOL_FONT_FLOOR and not _fits_one_line(need, width_pt, f):
            f -= 0.5
        return f
    disp = _disp_w(lines[0])
    f = COL1_FONT_MAX
    while f > SUBCOL_FONT_FLOOR and not _fits_one_line(disp, width_pt, f):
        f -= 0.5
    if _fits_one_line(disp, width_pt, f):
        return f
    f = COL1_FONT_MAX
    while f > SUBCOL_FONT_FLOOR and not _fits_two_lines(disp, width_pt, f):
        f -= 0.5
    return f


def tidy_label_cells(table, nrows):
    """Shrink col0/col1 inner margins and re-wrap their labels to <=2 balanced
    lines so a long 평가 항목 can't prise the row open one glyph at a time. Cells
    that span several columns (e.g. 제품 일정) use their full merged width so they
    do not wrap when there is plenty of room."""
    cw = E.TABLE_GEOMETRY["column_widths"]
    for r in range(1, nrows):
        for c in (0, 1):
            cell = table.cell(r, c)
            try:
                if cell.is_spanned:
                    continue
            except Exception:
                pass
            gs = int(cell._tc.get("gridSpan") or 1)
            span_w = sum(cw[c:c + gs])
            usable = (span_w * 0.75 - 2 * CELL_MARGIN_PT) / CHAR_PT
            cell.margin_left = Pt(CELL_MARGIN_PT)
            cell.margin_right = Pt(CELL_MARGIN_PT)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            txt = cell.text_frame.text
            if txt.strip():
                if c == 1 and gs == 1:                # item name: keep user breaks, never add a 3rd line
                    font = _fit_subcolumn_font(txt, span_w)
                    if "\n" in txt:                   # user already broke it -> keep those lines as-is
                        wrapped = txt
                    else:
                        cap = _line_cap_units(span_w, font) * SUBCOL_FIT_SAFETY
                        wrapped = txt if _disp_w(txt) <= cap else wrap_balanced(txt, cap)
                    rh = (table.rows[r].height or 0) / UNIT     # shrink so the name fits the row height
                    nlines = wrapped.count("\n") + 1
                    if rh > 0:
                        font = min(font, max(SUBCOL_FONT_FLOOR, (rh - 6.0) / (nlines * 1.45)))
                    E.set_cell_text(cell, wrapped, bold=True, size_pt=font)
                else:                                 # block title (col0) / merged cell
                    if "\n" in txt:                   # respect explicit breaks: keep lines, shrink to fit width
                        need = max((_disp_w(ln) for ln in txt.split("\n")), default=0)
                        f = LABEL_PT
                        while f > 6.5 and not _fits_one_line(need, span_w, f):
                            f -= 0.5
                        E.set_cell_text(cell, txt, bold=True, size_pt=f)
                    else:
                        wrapped = wrap_balanced(txt, usable)
                        E.set_cell_text(cell, wrapped, bold=True, size_pt=LABEL_PT)
                for para in cell.text_frame.paragraphs:
                    para.alignment = PP_ALIGN.CENTER


def _status_cell(cell, text, fill_hex, text_hex):
    E.set_cell_text(cell, text, bold=True, size_pt=STATUS_PT)
    cell.fill.background()
    for para in cell.text_frame.paragraphs:
        for run in para.runs:
            run.font.color.rgb = RGBColor.from_string(text_hex)


def harden_header_row(table):
    """The header row (year + month labels) must NEVER wrap: a wrapped header doubles the row
    height and shifts the whole timeline out of alignment. Force single-line cells, trim the
    inner margins, and shrink the font to fit the column width (size may drop, never wraps)."""
    cw = E.TABLE_GEOMETRY["column_widths"]
    for c in range(len(table.columns)):
        cell = table.cell(0, c)
        try:
            if cell.is_spanned:
                continue
        except Exception:
            pass
        tf = cell.text_frame
        tf.word_wrap = False
        cell.margin_left = Pt(1.0)
        cell.margin_right = Pt(1.0)
        txt = tf.text.strip()
        if not txt:
            continue
        gs = int(cell._tc.get("gridSpan") or 1)
        usable = max(8.0, sum(cw[c:c + gs]) - 4.0)        # column width minus the trimmed margins
        f = 12.0
        while f > 6.0 and _label_width_for(txt, f) > usable:
            f -= 0.5
        for para in tf.paragraphs:
            for run in para.runs:
                run.font.size = Pt(f)


def fill_status(table, statuses, nrows):
    E.set_cell_text(table.cell(0, 2), "현황", bold=True, size_pt=12)
    for r, st in statuses.items():
        if st:
            _status_cell(table.cell(r, 2), st, *STATUS_STYLE[st])
    # col0 (block-title) font is owned by tidy_label_cells (it shrinks pre-broken titles to fit);
    # do NOT force it back to LABEL_PT here or a 2-line title re-wraps into 3 lines.


# --------------------------------------------------------------------------- #
# Marker lanes (every evaluation item + the product row, placed at true date x)
# --------------------------------------------------------------------------- #
def recolor(sp, hl):
    if hl:
        sp.fill.solid()
        sp.fill.fore_color.rgb = HL_RGB[hl]


def center_group_in_row(shapes, row_top, row_h):
    """Shift a lane (circles + labels + arrow) as one block so its centre matches
    the table-row centre, regardless of the row's height."""
    shapes = [x for x in shapes if x is not None]
    if not shapes:
        return
    g_top = min(x.top for x in shapes)
    g_bot = max(x.top + x.height for x in shapes)
    delta = (row_top + (row_h - (g_bot - g_top)) / 2.0) - g_top
    for x in shapes:
        x.top = x.top + delta


def _arrow_source(resolver):
    arrows = sorted(resolver.by_name.get(E.ARROW_NAME, []), key=lambda s: s.top or 0)
    return E.S(arrows[0]) if arrows else None


def _send_behind(sp, ref_sp):
    """Move sp's element to just before ref_sp's in the spTree, so sp draws BEHIND ref_sp."""
    el = sp.sp._element
    el.getparent().remove(el)
    ref_sp.sp._element.addprevious(el)


def place_item_lane(slide, resolver, name, events, row_top, row_h, scale, warnings):
    """One evaluation item: clone a circle+label per check-post at its TRUE date x
    (override-aware, so identical dates line up vertically and stay date-ordered),
    colour it, span the arrow, then centre the whole lane in its row."""
    if not events:
        return
    centers = E.centers_for_items(events, "item", scale, warnings)
    locked_set = (scale.get("date_center_overrides", {}) or {}).get("locked_serials", set())
    vs = scale.get("lane_vscale", 1.0)               # <1.0 shrinks the lane when rows are crowded
    csz = 30.0 * vs
    circle_top = row_top + 1.0
    label_top = circle_top + csz
    circ_src = resolver.by_name_one(GEN_CIRCLE)
    lbl_src = resolver.by_name_one(GEN_LABEL)
    fs = scale.get("label_font_scale", 1.0)
    placed, shapes = [], []
    for idx, item in enumerate(events):
        cir = E.clone_shape(slide, circ_src, f"{name} circle {idx + 1}", circle_top)
        lbl = E.clone_shape(slide, lbl_src, f"{name} label {idx + 1}", label_top)
        E.place_pair_at_center(lbl, cir, item, centers[idx], scale, {
            "circle_top": circle_top, "label_top": label_top,
            "lock_circle_center": True, "lane": "item"})   # circle stays exactly on its date x
        if vs != 1.0:                                   # crowded lanes: shrink circle to fit row
            cc = cir.left + cir.width / 2.0
            cir.width = csz
            cir.height = csz
            cir.left = cc - csz / 2.0
            cir.top = circle_top
            for para in cir.sp.text_frame.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(14.67 * vs)
        _center_label_on_circle(lbl, cir, item["label"], fs * vs)   # label centred on circle (dx == 0)
        recolor(cir.sp, item.get("hl"))
        placed.append({"label": lbl, "circle": cir, "item": item})
        shapes += [cir, lbl]
    # one connector arrow PER segment: a blank check-post row splits the lane, so the row
    # before the gap ends one arrow and the row after starts a new one.
    seg_centers = {}
    for p in placed:
        seg_centers.setdefault(p["item"].get("seg", 0), []).append(
            p["circle"].left + p["circle"].width / 2.0)
    arrows = []
    for si, cs in sorted(seg_centers.items()):
        if len(cs) >= 2:
            arrow = E.clone_shape(slide, _arrow_source(resolver), f"{name} arrow {si}", circle_top + 1.9)
            if vs != 1.0:
                arrow.top = circle_top + csz / 2.0 - arrow.height / 2.0   # centre on shrunk circle
            E.move_lane_arrow(arrow, cs, "lane1")
            arrows.append(arrow)
    if len(placed) == 1:                              # single milestone: arrow from the timeline
        mc = placed[0]["circle"].left + placed[0]["circle"].width / 2.0   # start (right of 현황) up
        fl = E.timeline_frame_left()                  # to that one marker -- it is the terminus
        arrow = E.clone_shape(slide, _arrow_source(resolver), f"{name} arrow 0", circle_top + 1.9)
        if vs != 1.0:
            arrow.top = circle_top + csz / 2.0 - arrow.height / 2.0
        arrow.left = fl + 2.0
        arrow.width = max(20.0, mc - csz / 2.0 - (fl + 2.0))
        arrows.append(arrow)
    # draw the connector arrows BEHIND this lane's circles so they never cover a marker's number
    first_circle = placed[0]["circle"] if placed else None
    for arrow in arrows:
        shapes.append(arrow)
        if first_circle is not None:
            _send_behind(arrow, first_circle)
    center_group_in_row(shapes, row_top, row_h)


def deconflict_labels_fixed_circles(placed, lo=0.0, hi=1280.0, gap=4.0):
    """Keep every circle at its true date x; place each label as close to centred under its
    own circle as possible. Non-colliding labels stay exactly centred (dx~0); colliding
    labels are spread symmetrically around their cluster with the minimum total displacement
    (isotonic / pool-adjacent-violators), and the circles are never moved -- so markers keep
    their correct month column and date order while labels stop piling up on one side."""
    if not placed:
        return
    its = []
    for p in placed:
        cir, lbl = p["circle"], p["label"]
        c = (cir.left + cir.width / 2.0) if cir is not None else (lbl.left + lbl.width / 2.0)
        sw = p.get("spacing_w") or lbl.width     # space by glyph width; padded boxes may overlap
        its.append({"lbl": lbl, "t": c, "w": lbl.width, "sw": sw})
    its.sort(key=lambda e: e["t"])
    n = len(its)
    if n == 1:
        its[0]["lbl"].left = its[0]["t"] - its[0]["w"] / 2.0
        return
    # minimum centre-to-centre spacing so two label TEXTS don't overlap
    msep = [(its[i]["sw"] + its[i + 1]["sw"]) / 2.0 + gap for i in range(n - 1)]
    pref = [0.0] * n
    for i in range(1, n):
        pref[i] = pref[i - 1] + msep[i - 1]
    # the ordered-spacing constraint c[i+1]-c[i] >= msep[i] becomes u[i+1] >= u[i] for
    # u[i] = c[i] - pref[i]; the L2-optimal monotone u is the isotonic regression of y.
    y = [its[i]["t"] - pref[i] for i in range(n)]
    vals, wts, cnt = [], [], []
    for yi in y:
        vals.append(yi); wts.append(1.0); cnt.append(1)
        while len(vals) > 1 and vals[-2] > vals[-1]:
            v2, w2, k2 = vals.pop(), wts.pop(), cnt.pop()
            v1, w1, k1 = vals.pop(), wts.pop(), cnt.pop()
            vals.append((v1 * w1 + v2 * w2) / (w1 + w2)); wts.append(w1 + w2); cnt.append(k1 + k2)
    g = []
    for v, k in zip(vals, cnt):
        g.extend([v] * k)
    c = [g[i] + pref[i] for i in range(n)]
    # slide bounds: pull in only an overflowing tail/head (isolated labels untouched)
    hib = hi - its[n - 1]["w"] / 2.0
    if c[n - 1] > hib:
        c[n - 1] = hib
        for i in range(n - 2, -1, -1):
            c[i] = min(c[i], c[i + 1] - msep[i])
    lob = lo + its[0]["w"] / 2.0
    if c[0] < lob:
        c[0] = lob
        for i in range(1, n):
            c[i] = max(c[i], c[i - 1] + msep[i - 1])
    for i in range(n):
        its[i]["lbl"].left = c[i] - its[i]["w"] / 2.0


def place_product_lane(slide, resolver, items, scale, warnings):
    """제품 일정 master lane at true date x, cloning extra circle/label pairs when
    there are more products than the template's pre-built slots."""
    tmpl = E.TEMPLATE_NAMES["product"]
    # Drop the template's pre-built product slots that this input doesn't use,
    # otherwise stale sample markers (e.g. Post4 / 완료) are left on the slide.
    if len(items) < len(tmpl):
        leftover = {n for pair in tmpl[len(items):] for n in pair}
        for sp in list(slide.shapes):
            if sp.name in leftover:
                sp._element.getparent().remove(sp._element)
    pairs = []
    for i in range(len(items)):
        if i < len(tmpl):
            ln, cn = tmpl[i]
            pairs.append((resolver.by_name_one(ln), resolver.by_name_one(cn)))
        else:
            sln, scn = tmpl[-1]
            src_lbl = resolver.by_name_one(sln)
            src_cir = resolver.by_name_one(scn)
            cir = E.clone_shape(slide, src_cir, f"product extra circle {i + 1}", src_cir.top)
            lbl = E.clone_shape(slide, src_lbl, f"product extra label {i + 1}", src_lbl.top)
            pairs.append((lbl, cir))
    centers = E.centers_for_items(items, "product", scale, warnings)
    fs = scale.get("label_font_scale", 1.0)
    for (lbl, cir), item, c in zip(pairs, items, centers):
        E.place_pair_at_center(lbl, cir, item, c, scale, {
            "lock_circle_center": True, "lane": "product"})   # circle stays exactly on its date x
        _center_label_on_circle(lbl, cir, item["label"], fs)   # label centred on circle (dx == 0)
    if items:                                  # paint the final product milestone dark green
        last_cir = pairs[len(items) - 1][1].sp
        last_cir.fill.solid()
        last_cir.fill.fore_color.rgb = PRODUCT_LAST_RGB
    rc = None                                  # rightmost product marker = the lane's terminus
    for (lbl, cir), item in zip(pairs, items):
        c = cir.left + cir.width / 2.0
        if rc is None or c > rc[0]:
            rc = (c, cir.width)
    return rc


def remove_unused_template_markers(slide):
    """Every evaluation lane is cloned from the generic source, so drop the
    template's leftover pre-built marker shapes (and their arrows) to avoid stale
    text from the original sample."""
    names = set()
    for ln, cn in E.TEMPLATE_NAMES["lane1"] + E.TEMPLATE_NAMES["lane2"]:
        names.add(ln)
        names.add(cn)
    for sp in list(slide.shapes):
        if sp.name in names or sp.name == E.ARROW_NAME:
            sp._element.getparent().remove(sp._element)


# --------------------------------------------------------------------------- #
# Table rows (one per evaluation item; block title spans its items in col0)
# --------------------------------------------------------------------------- #
def build_item_rows(table, blocks, row_h):
    """Rebuild the evaluation area: product row stays, then one row per item.
    col0 = block title (merged across the block's items), col1 = item name.
    Returns the 0-based table-row index of each item, block by block."""
    tbl = table._tbl
    trs = tbl.findall(qn("a:tr"))
    first_tpl = copy.deepcopy(trs[2])     # block-first item: col0 is a merge origin
    cont_tpl = copy.deepcopy(trs[3])      # continuation item: col0 carries vMerge
    for tr in trs[2:]:                    # drop all template eval rows
        tbl.remove(tr)
    for block in blocks:
        n = len(block["items"])
        tbl.append(copy.deepcopy(first_tpl))
        for _ in range(n - 1):
            tbl.append(copy.deepcopy(cont_tpl))

    rows_of_block = []
    trs = tbl.findall(qn("a:tr"))
    r = 2
    for block in blocks:
        n = len(block["items"])
        merged = (n == 1 and block["items"][0].get("merged"))
        tcs0 = trs[r].findall(qn("a:tc"))
        c0 = tcs0[0]
        if n > 1:
            c0.set("rowSpan", str(n))
        elif c0.get("rowSpan") is not None:
            del c0.attrib["rowSpan"]
        title0 = str(block["title"]).strip() or block["title"]   # keep '(...)': no paren-stripping
        E.set_cell_text(table.cell(r, 0), title0, bold=True)
        if merged:                                  # 상위 Plan만 -> col0+col1 한 칸으로 병합
            E.set_cell_text(table.cell(r, 1), "")    # remove template name before merging
            c0.set("gridSpan", "2")
            tcs0[1].set("hMerge", "1")
        idxs = []
        for i, item in enumerate(block["items"]):
            rr = r + i
            table.rows[rr].height = Emu(int(round(row_h * UNIT)))
            if not merged:
                E.set_cell_text(table.cell(rr, 1), item["name"], bold=True,
                                size_pt=E._subcolumn_size([item["name"]]))
            idxs.append(rr)
        rows_of_block.append(idxs)
        r += n
    return rows_of_block


def update_timeline_header(table, scale):
    set_text = E.set_cell_text
    set_text(table.cell(0, 0), E.year_range_label(scale["start_serial"], scale["end_serial"]))
    # 현재(올해) 연도는 월만, 과거(및 다른) 연도는 'YY.M월'로 연도까지 표기
    base_year = E.excel_serial_to_date(E.today_serial()).year
    selected = scale["header_month_keys"]
    for index, column in enumerate(E.timeline_month_columns()):
        key = selected[index] if index < len(selected) else None
        text = "" if key is None else E.month_label(E.serial_from_month_key(key), base_year)
        set_text(table.cell(0, column["cell_index"]), text)


def fit_month_columns(table, n_active):
    """Make the table hold exactly n_active month columns: drop the template's
    leftover months when fewer are needed, or CLONE extra month columns when the
    timeline spans more than the template's pre-built 8 (e.g. >8 active months).
    Widths are then re-applied from the geometry."""
    keep = 3 + n_active                   # 평가항목 + 하위 + 현황 + shown months
    tbl = table._tbl
    grid = tbl.find(qn("a:tblGrid"))
    cur = len(grid.findall(qn("a:gridCol")))
    if cur > keep:                        # too many -> drop the trailing template months
        for gc in grid.findall(qn("a:gridCol"))[keep:]:
            grid.remove(gc)
        for tr in tbl.findall(qn("a:tr")):
            for tc in tr.findall(qn("a:tc"))[keep:]:
                tr.remove(tc)
    elif cur < keep:                      # too few -> clone the last month column
        add = keep - cur
        for _ in range(add):
            grid.append(copy.deepcopy(grid.findall(qn("a:gridCol"))[-1]))
        for tr in tbl.findall(qn("a:tr")):
            ref = tr.findall(qn("a:tc"))[-1]
            for _ in range(add):
                ntc = copy.deepcopy(ref)
                for a in ("gridSpan", "rowSpan", "hMerge", "vMerge"):
                    if ntc.get(a) is not None:
                        del ntc.attrib[a]
                for t in ntc.iter(qn("a:t")):
                    t.text = ""
                tr.append(ntc)
    for gc, w in zip(grid.findall(qn("a:gridCol")), E.TABLE_GEOMETRY["column_widths"]):
        gc.set("w", str(int(round(w * UNIT))))
    _unify_month_separators(tbl)
    _blank_right_border(tbl)


def _unify_month_separators(tbl):
    """Unify the internal vertical separators of all month columns. The template's last month
    column is the table's outer edge, so its right border is noFill; when more than 8 months are
    needed that column gets cloned, which drops/recolours the internal separators. Restore the
    right border of every month column except the last to the standard interior line actually
    used in the table (the outer-right edge is left to the following _blank_right_border)."""
    month_start = E.TABLE_GEOMETRY.get("month_start_column", 3)
    ref = None
    for tr in tbl.findall(qn("a:tr")):
        for tc in tr.findall(qn("a:tc")):
            pr = tc.find(qn("a:tcPr"))
            if pr is None:
                continue
            ln = pr.find(qn("a:lnR"))
            if ln is not None and ln.find(qn("a:solidFill")) is not None:
                ref = ln
                break
        if ref is not None:
            break
    if ref is None:
        return
    for tr in tbl.findall(qn("a:tr")):
        tcs = tr.findall(qn("a:tc"))
        for c in range(month_start, len(tcs) - 1):      # exclude the last (outer-edge) column
            pr = tcs[c].find(qn("a:tcPr"))
            if pr is None:
                continue
            old = pr.find(qn("a:lnR"))
            new = copy.deepcopy(ref)
            if old is not None:
                pr.replace(old, new)
            else:
                lnl = pr.find(qn("a:lnL"))
                if lnl is not None:
                    lnl.addnext(new)
                else:
                    pr.insert(0, new)


def _blank_right_border(tbl):
    """Make the table's right-most edge borderless (a deleted dummy column used to
    sit there; the new last column must not show an outer line)."""
    for tr in tbl.findall(qn("a:tr")):
        tcs = tr.findall(qn("a:tc"))
        if not tcs:
            continue
        tcPr = tcs[-1].find(qn("a:tcPr"))
        if tcPr is None:
            continue
        lnR = tcPr.find(qn("a:lnR"))
        if lnR is None:
            continue
        for tag in ("a:solidFill", "a:gradFill", "a:pattFill", "a:noFill"):
            for e in lnR.findall(qn(tag)):
                lnR.remove(e)
        lnR.insert(0, lnR.makeelement(qn("a:noFill"), {}))


def reposition_now(resolver, scale, today):
    x = E.x_for_date(today, scale, [], warn=False)
    resolver.by_name_one(E.TEMPLATE_NAMES["now_marker"]).left = x - \
        resolver.by_name_one(E.TEMPLATE_NAMES["now_marker"]).width / 2
    resolver.by_name_one(E.TEMPLATE_NAMES["now_label"]).left = x + 12.1


# --------------------------------------------------------------------------- #
# Native-table normalisation (keeps appearance; restores hand-editability)
# --------------------------------------------------------------------------- #
def rebuild_table_native(slide, gf_shape):
    """Rebuild the schedule table as a fresh python-pptx (native) table, copying
    every cell's formatting verbatim so the appearance is pixel-identical.

    Why: the table is produced by raw-XML surgery on the template table (insert
    the 현황 column, delete unused month columns, append evaluation rows, retrofit
    row/column merges). The result is schema-valid and renders correctly, but
    PowerPoint's table editor does NOT treat the grafted cells as native cells --
    typing/deleting lands in the wrong cell, columns can't be removed, and undo
    blanks the table. Re-creating the grid through add_table() restores normal
    editing; transplanting each cell's <a:tcPr>/<a:txBody> keeps the look intact.
    """
    old_gf = gf_shape._element
    tbl = gf_shape.table._tbl
    grid = tbl.find(qn("a:tblGrid"))
    widths = [int(c.get("w") or 0) for c in grid.findall(qn("a:gridCol"))]
    trs = tbl.findall(qn("a:tr"))
    heights = [int(tr.get("h") or 0) for tr in trs]
    nrow, ncol = len(trs), len(widths)

    # snapshot every cell's formatting + content, and detect merge origins
    cell_xml = {}
    merges = []
    for r, tr in enumerate(trs):
        for c, tc in enumerate(tr.findall(qn("a:tc"))):
            cell_xml[(r, c)] = (copy.deepcopy(tc.find(qn("a:tcPr"))),
                                copy.deepcopy(tc.find(qn("a:txBody"))))
            gs = int(tc.get("gridSpan") or 1)
            rs = int(tc.get("rowSpan") or 1)
            if (gs > 1 or rs > 1) and not tc.get("hMerge") and not tc.get("vMerge"):
                merges.append((r, c, r + rs - 1, c + gs - 1))

    left, top, width, height = gf_shape.left, gf_shape.top, gf_shape.width, gf_shape.height

    new_gf = slide.shapes.add_table(nrow, ncol, left, top, width, height)
    ntbl = new_gf.table
    for i, w in enumerate(widths):
        ntbl.columns[i].width = Emu(w)
    for i, h in enumerate(heights):
        ntbl.rows[i].height = Emu(h)

    # transplant tcPr + txBody into every cell of the fresh native grid
    for r, tr in enumerate(ntbl._tbl.findall(qn("a:tr"))):
        for c, tc in enumerate(tr.findall(qn("a:tc"))):
            src_pr, src_tx = cell_xml[(r, c)]
            if src_tx is not None:
                tc.replace(tc.find(qn("a:txBody")), copy.deepcopy(src_tx))
            old_pr = tc.find(qn("a:tcPr"))
            if src_pr is not None:
                if old_pr is not None:
                    tc.replace(old_pr, copy.deepcopy(src_pr))
                else:
                    tc.append(copy.deepcopy(src_pr))

    # re-apply merges through the API (produces canonical merge XML)
    for r0, c0, r1, c1 in merges:
        ntbl.cell(r0, c0).merge(ntbl.cell(r1, c1))

    # explicit per-cell borders must stay dominant -> "No Style, No Grid"
    tblPr = ntbl._tbl.find(qn("a:tblPr"))
    sid = tblPr.find(qn("a:tableStyleId")) if tblPr is not None else None
    if sid is not None:
        sid.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"

    # restore z-order: drop the old frame in, take its slot, remove it
    parent = old_gf.getparent()
    parent.insert(list(parent).index(old_gf), new_gf._element)
    parent.remove(old_gf)


def strip_stale_thumbnail(pptx_path):
    """Remove the template's embedded preview image (docProps/thumbnail.*) that
    python-pptx carries over from the template. Otherwise Explorer/Finder shows
    the *template's* slide as this file's thumbnail. PowerPoint regenerates a
    correct thumbnail the next time the file is opened and saved.
    """
    import zipfile
    import re
    src = str(pptx_path)
    with zipfile.ZipFile(src, "r") as zin:
        names = zin.namelist()
        if not any(n.lower().startswith("docprops/thumbnail") for n in names):
            return
        items = [(i, zin.read(i.filename)) for i in zin.infolist()]
    tmp = src + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for info, data in items:
            name = info.filename
            if name.lower().startswith("docprops/thumbnail"):
                continue
            if name == "_rels/.rels":
                text = data.decode("utf-8")
                text = re.sub(r'<Relationship[^>]*?thumbnail[^>]*?/>', '', text)
                data = text.encode("utf-8")
            zout.writestr(info, data)
    os.replace(tmp, src)


# --------------------------------------------------------------------------- #
# Main
def _shape_by_name(slide, name):
    for sp in slide.shapes:
        if sp.name == name:
            return sp
    return None


def _subtitle_shape(slide, table_sp):
    """The 부제 shape. Prefer the templated name; else the header-band shape whose
    bottom is closest above the table (works whether 부제 is a textbox or a
    corporate placeholder)."""
    named = _shape_by_name(slide, E.TEMPLATE_NAMES["punchline"])
    if named is not None and named.top is not None:
        return named
    best, best_bottom = None, None
    for sp in slide.shapes:
        if sp is table_sp or sp.top is None or sp.height is None:
            continue
        if sp.top < (table_sp.top or 0):                 # header band, above the table
            bottom = sp.top + sp.height
            if best_bottom is None or bottom > best_bottom:
                best, best_bottom = sp, bottom
    return best


def reframe_chart(prs, slide):
    """Move the table + every milestone shape as ONE rigid group so the table is
    horizontally centred on the slide and starts a fixed gap below the 부제.
    제목/부제/꼬리말(placeholder)은 제자리에 둔다. The rigid shift preserves all
    table<->marker alignment (they were already mutually aligned)."""
    table_sp = next((sp for sp in slide.shapes if sp.has_table), None)
    if table_sp is None or table_sp.left is None or table_sp.top is None:
        return

    # true rendered width = sum of column widths; keep the frame box in sync
    real_w = sum(int(c.width or 0) for c in table_sp.table.columns)
    if real_w > 0:
        table_sp.width = Emu(real_w)
    tw = int(table_sp.width or 0)

    desired_left = int(round((int(prs.slide_width) - tw) / 2.0))     # exact horizontal centre
    sub = _subtitle_shape(slide, table_sp)
    if sub is not None and sub.top is not None and sub.height is not None:
        desired_top = int(sub.top + sub.height + round(GAP_BELOW_SUBTITLE * UNIT))
    else:
        desired_top = int(table_sp.top)

    dx = desired_left - int(table_sp.left)
    dy = desired_top - int(table_sp.top)
    if dx == 0 and dy == 0:
        return

    keep = {E.TEMPLATE_NAMES["title"]}              # 제목 고정
    if sub is not None:
        keep.add(sub.name)                          # 부제 고정
    for sp in slide.shapes:
        if sp.is_placeholder or sp.name in keep:    # 제목/부제/꼬리말(placeholder)은 제자리
            continue
        if sp.left is not None:
            sp.left = Emu(int(sp.left) + dx)
        if sp.top is not None:
            sp.top = Emu(int(sp.top) + dy)


def _set_white_bg(slide):
    """Force a solid-white slide background (overrides any inherited fill)."""
    cSld = slide._element.find(qn("p:cSld"))
    old = cSld.find(qn("p:bg"))
    if old is not None:
        cSld.remove(old)
    bg = parse_xml(
        '<p:bg xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
        '<a:effectLst/></p:bgPr></p:bg>')
    cSld.insert(0, bg)            # <p:bg> must be the first child of <p:cSld>


def _apply_background(slide, mode):
    """'memory' keeps the template's layout decoration (사내 배경 그대로).
    'none' strips the layout's non-placeholder shapes (logo/watermark/footer
    images, lines, labels) and forces white -> 표/부제/마일스톤만 남긴다."""
    if str(mode).lower() != "none":
        return
    lay = slide.slide_layout
    spTree = lay._element.find(qn("p:cSld")).find(qn("p:spTree"))
    for sh in list(lay.shapes):
        if not sh.is_placeholder:                 # keep placeholders, drop decoration
            spTree.remove(sh._element)
    _set_white_bg(slide)


def _whiten_footer_overlap(prs, slide):
    """When the table runs down into the layout footer band (e.g. 장식 배경이 있는 템플릿의
    '함께 성장하는...' 푸터), the transparent cells let that text show through and hurt
    readability. Paint ONLY the cells that overlap a footer text box solid white.
    The big watermark image (no text) is ignored, so the rest of the background stays."""
    table_sp = next((s for s in slide.shapes if s.has_table), None)
    if table_sp is None or table_sp.left is None or table_sp.top is None:
        return
    SH = int(prs.slide_height)
    foots = []
    for sh in slide.slide_layout.shapes:
        if sh.top is None or sh.height is None or sh.width is None:
            continue
        has_text = sh.has_text_frame and sh.text_frame.text.strip()
        if has_text and sh.top > SH * 0.85:          # footer text in the bottom band
            foots.append((int(sh.left), int(sh.top), int(sh.left + sh.width), int(sh.top + sh.height)))
    if not foots:
        return
    tbl = table_sp.table
    xs = [int(table_sp.left)]
    for c in tbl.columns:
        xs.append(xs[-1] + int(c.width or 0))
    ys = [int(table_sp.top)]
    for r in tbl.rows:
        ys.append(ys[-1] + int(r.height or 0))
    white = RGBColor(0xFF, 0xFF, 0xFF)
    hit_rows = set()
    for ri in range(len(tbl.rows)):
        for ci in range(len(tbl.columns)):
            cl, ct, cr, cb = xs[ci], ys[ri], xs[ci + 1], ys[ri + 1]
            if any(not (cr <= fl or cl >= fr or cb <= ft or ct >= fb) for fl, ft, fr, fb in foots):
                hit_rows.add(ri)
                break
    for ri in hit_rows:                                # 겹치는 셀이 있는 행은 전체를 흰색으로
        for ci in range(len(tbl.columns)):
            cell = tbl.cell(ri, ci)
            cell.fill.solid()
            cell.fill.fore_color.rgb = white


# --------------------------------------------------------------------------- #
# Drill-down: a detail slide per evaluation item that has a '<item> 상세' sheet
# --------------------------------------------------------------------------- #
def _detail_sheet_name(item_name):
    """Sheet name a detail tab has for an evaluation item -- MUST match the
    Module_Drilldown VBA: invalid chars -> '-', clipped to 26 chars, + ' 상세'."""
    s = str(item_name)
    for ch in "[]:*?/\\":
        s = s.replace(ch, "-")
    s = s.strip()[:26].strip()
    return (s + " 상세") if s else ""


def _list_sheets(path):
    """Worksheet names (openpyxl; Excel-COM fallback for DRM-wrapped files)."""
    try:
        return list(openpyxl.load_workbook(path, read_only=True).sheetnames)
    except Exception:
        pass
    try:
        import comtypes.client
        excel = comtypes.client.CreateObject("Excel.Application", dynamic=True)
        try:
            for prop, val in (("Visible", False), ("DisplayAlerts", False)):
                try:
                    setattr(excel, prop, val)
                except Exception:
                    pass
            wb = excel.Workbooks.Open(os.path.abspath(path), 0, True)
            try:
                cnt = int(wb.Worksheets.Count)
                return [str(wb.Worksheets(i + 1).Name) for i in range(cnt)]
            finally:
                wb.Close(False)
        finally:
            excel.Quit()
    except Exception:
        return []


def _copy_slide_into(dst_prs, src_slide):
    """Append a copy of src_slide to dst_prs. Both decks use the same template, so the
    layout is shared; we copy every shape element (autoshapes/table/textboxes have no
    external rels -- background images live on the shared layout)."""
    layout = dst_prs.slides[0].slide_layout
    new = dst_prs.slides.add_slide(layout)
    for sh in list(new.shapes):                       # drop the layout's inherited placeholders
        sh._element.getparent().remove(sh._element)
    for sh in src_slide.shapes:
        new.shapes._spTree.append(copy.deepcopy(sh._element))
    return new


def _link_shape(sp, host_part, target_slide):
    """Whole-shape click action -> jump to target_slide (slideshow)."""
    rId = host_part.relate_to(target_slide.part, RT.SLIDE)
    cNvPr = sp._element.find(qn("p:nvSpPr")).find(qn("p:cNvPr"))
    old = cNvPr.find(qn("a:hlinkClick"))
    if old is not None:
        cNvPr.remove(old)
    cNvPr.insert(0, cNvPr.makeelement(qn("a:hlinkClick"),
                                      {qn("r:id"): rId, "action": "ppaction://hlinksldjump"}))


def _link_run(run, host_part, target_slide):
    """Text-run click action -> jump to target_slide (slideshow)."""
    rId = host_part.relate_to(target_slide.part, RT.SLIDE)
    rPr = run._r.get_or_add_rPr()
    old = rPr.find(qn("a:hlinkClick"))
    if old is not None:
        rPr.remove(old)
    rPr.append(rPr.makeelement(qn("a:hlinkClick"),
                               {qn("r:id"): rId, "action": "ppaction://hlinksldjump"}))


def _link_item_to_slide(main_slide, item_name, target_slide):
    """Find the table cell whose text == item_name, mark it as a link (blue + underline)
    and attach a slide-jump to its runs. Returns True if the cell was found."""
    want = str(item_name).replace("\n", " ").strip()
    table_sp = next((s for s in main_slide.shapes if s.has_table), None)
    if table_sp is None:
        return False
    for row in table_sp.table.rows:
        for cell in row.cells:
            if cell.text.replace("\n", " ").strip() == want:
                for para in cell.text_frame.paragraphs:
                    for run in para.runs:
                        run.font.underline = True
                        run.font.color.rgb = RGBColor(0x14, 0x28, 0xA0)
                        _link_run(run, main_slide.part, target_slide)
                return True
    return False


def _add_home_button(det_slide, target_slide, title_name=None):
    """Small house icon (roof triangle + body) just to the right of the 표제, linked
    back to target_slide (the master). Two shapes both carry the click action."""
    title_name = title_name or E.TEMPLATE_NAMES["title"]
    title = next((s for s in det_slide.shapes if s.name == title_name), None)
    if title is not None and title.left is not None:
        try:
            sz = title.text_frame.paragraphs[0].runs[0].font.size
            fpt = sz.pt if sz else 28.0
        except Exception:
            fpt = 28.0
        est = _disp_w(title.text_frame.text.strip()) * (fpt / 0.75) * 0.62
        x = int(title.left) + int(round((est + 24) * UNIT))
        cy = int(title.top or 0) + int(title.height or 0) // 2
    else:
        x = int(round(300 * UNIT))
        cy = int(round(60 * UNIT))
    P = lambda v: int(round(v * UNIT))
    hw = 24.0
    top = cy - P(11)
    blue = RGBColor(0x14, 0x28, 0xA0)
    roof = det_slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Emu(x), Emu(top), Emu(P(hw)), Emu(P(11)))
    body = det_slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(x + P(5)), Emu(top + P(10)), Emu(P(hw - 10)), Emu(P(11)))
    for sp in (roof, body):
        sp.fill.solid()
        sp.fill.fore_color.rgb = blue
        sp.line.fill.background()
        sp.shadow.inherit = False
        _link_shape(sp, det_slide.part, target_slide)
    return roof


# --------------------------------------------------------------------------- #
def _build_one(data, template_path, background):
    """Render ONE milestone slide from a parsed `data` dict; return (prs, warnings, n_items)."""
    warnings = []
    today = E.today_serial()
    blocks = data["blocks"]

    all_dates = [it["date"] for it in data["product"] if E.is_number(it["date"])]
    for block in blocks:
        for item in block["items"]:
            all_dates += [e["date"] for e in item["events"] if E.is_number(e["date"])]
    if not all_dates:
        all_dates = [today]
    scale, weights = build_scale(all_dates)

    item_names = [it["name"] for b in blocks for it in b["items"]]
    col1_w = compute_col1_width(item_names)
    layout = compute_timeline_layout(data["product"], blocks, scale["header_month_keys"], col1_w)
    tx = install_geometry(layout["month_widths"], col1_w)
    # every date is locked to its computed x: circles never move off their date, identical
    # dates line up across lanes, and labels are simply centred on the circles afterwards.
    scale["date_center_overrides"] = {"centers": layout["centers"],
                                      "locked_serials": set(layout["centers"].keys())}
    scale["label_font_scale"] = layout["font_scale"]

    prs = Presentation(template_path)
    slide = prs.slides[0]
    resolver = E.Resolver(slide)
    table = E.get_table(resolver)
    insert_status_structure(table)

    resolver.by_name_one(E.TEMPLATE_NAMES["title"]).set_text(f" {data['title']}")
    resolver.by_name_one(E.TEMPLATE_NAMES["punchline"]).set_text(f" ■ {data['punchline']}")
    pt = data.get("product_title")          # A4 값이 있으면 제품행 라벨로(없으면 템플릿 '제품 일정' 유지)
    if pt and str(pt).strip():
        E.set_cell_text(table.cell(1, 0), str(pt).strip(), bold=True)   # keep '(...)': no paren-stripping

    # ---- evaluation rows: label-aware uniform row height (+ shrink lane if crowded) ----
    n_items = sum(len(b["items"]) for b in blocks)
    max_label_lines = 1
    max_name_lines = 1
    for b in blocks:
        for it in b["items"]:
            max_name_lines = max(max_name_lines, str(it.get("name", "")).count("\n") + 1)
            for e in it["events"]:
                max_label_lines = max(max_label_lines, str(e.get("label", "")).count("\n") + 1)
    lane_full = 30.0 + max_label_lines * 16.0       # circle(30) + stacked label lines(16/line)
    name_block = 8.0 + max_name_lines * (COL1_FONT_MAX * 1.45)   # col1 name height (line ~1.45*font)
    product_h = 76.0
    # budget rows against where the table ACTUALLY starts (reframe puts it below the 부제) so it
    # never runs past TABLE_BOTTOM into the slide edge when there are many or tall rows.
    tbl_shape = next((sp for sp in slide.shapes if sp.has_table), None)
    sub = _subtitle_shape(slide, tbl_shape) if tbl_shape is not None else None
    if sub is not None and sub.top is not None and sub.height is not None:
        eff_top = (int(sub.top) + int(sub.height)) / UNIT + GAP_BELOW_SUBTITLE
    else:
        eff_top = TABLE_TOP
    area_top = eff_top + HDR_H + product_h
    avail = TABLE_BOTTOM - area_top
    if n_items:
        fair = avail / n_items
        row_h = min(fair, max(ITEM_ROW_H, lane_full + 8.0, name_block))
        lane_vscale = 1.0 if row_h >= lane_full + 2.0 else max(0.62, (row_h - 2.0) / lane_full)
    else:
        row_h, lane_vscale = ITEM_ROW_H, 1.0
    if lane_vscale < 1.0:
        warnings.append(f"평가 항목이 많아({n_items}개) 마일스톤을 약 {int(round(lane_vscale * 100))}%로 축소 배치했습니다. "
                        "항목 분할(드릴다운)이나 Plan 분리를 권장합니다.")
    scale["lane_vscale"] = lane_vscale
    table.rows[1].height = Emu(int(round(product_h * UNIT)))
    rows_of_block = build_item_rows(table, blocks, row_h)

    # ---- product master lane ----
    prod_term = place_product_lane(slide, resolver, data["product"], scale, warnings)

    # ---- one lane per evaluation item, centred on its own row ----
    statuses = {}
    for bi, block in enumerate(blocks):
        for ii, item in enumerate(block["items"]):
            rr = rows_of_block[bi][ii]
            row_top = TABLE_TOP + sum((table.rows[k].height or 0) / UNIT for k in range(rr))
            place_item_lane(slide, resolver, f"b{bi}_i{ii}", item["events"],
                            row_top, row_h, scale, warnings)
            statuses[rr] = compute_status(item["events"], today)

    remove_unused_template_markers(slide)
    fit_month_columns(table, len(scale["header_month_keys"]))
    update_timeline_header(table, scale)
    reposition_now(resolver, scale, today)

    # static product connector arrow -> span the (rescaled) timeline frame
    for sp in slide.shapes:
        if sp.name == "오른쪽 화살표 134":
            pa = E.S(sp)
            l0, r0 = pa.left, pa.left + pa.width
            pa.left = tx(l0)
            if prod_term is not None:                  # end at the LAST product marker (its terminus),
                end = prod_term[0] - prod_term[1] / 2.0  # not the frame edge -- so it never overruns
                pa.width = max(20.0, end - pa.left)
            else:
                pa.width = tx(r0) - tx(l0)
            break

    nrows = len(table._tbl.findall(qn("a:tr")))
    tidy_label_cells(table, nrows)
    fill_status(table, statuses, nrows)
    harden_header_row(table)

    # Final step: re-emit the table as a native PowerPoint table so it stays
    # fully editable by hand (same appearance). See rebuild_table_native().
    for shp in list(slide.shapes):
        if shp.has_table:
            rebuild_table_native(slide, shp)
            break

    # final layout: centre the table horizontally + start it below the 부제
    reframe_chart(prs, slide)
    _apply_background(slide, background)          # 'memory' 배경 유지 / 'none' 배경 제거
    _whiten_footer_overlap(prs, slide)           # 푸터 텍스트와 겹치는 셀만 흰 배경 처리
    return prs, warnings, n_items


# --------------------------------------------------------------------------- #
def generate(input_path, template_path, output_path, background="none"):
    """Main milestone slide + a drill-down detail slide for every evaluation item that
    has a '<item> 상세' sheet. Click the item (blue, underlined) -> its detail slide;
    the house icon by the 표제 -> back to the master slide. (slideshow / Ctrl+click)"""
    main = parse_input(input_path)
    prs, warnings, n_items = _build_one(main, template_path, background)
    main_slide = prs.slides[0]
    sheets = _list_sheets(input_path)
    n_det = 0
    for b in main["blocks"]:
        for it in b["items"]:
            ds = _detail_sheet_name(it["name"])
            if ds and ds in sheets:
                ddata = parse_input(input_path, sheet=ds)
                dprs, dwarn, _ = _build_one(ddata, template_path, background)
                warnings += dwarn
                det = _copy_slide_into(prs, dprs.slides[0])
                _add_home_button(det, main_slide)
                if _link_item_to_slide(main_slide, it["name"], det):
                    n_det += 1
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))
    strip_stale_thumbnail(out)
    return {"output": str(out), "warnings": warnings, "detail_slides": n_det,
            "counts": {"product": len(main["product"]), "blocks": len(main["blocks"]),
                       "items": n_items, "detail_slides": n_det,
                       "checkposts": sum(len(it["events"]) for b in main["blocks"] for it in b["items"])}}


if __name__ == "__main__":
    import argparse
    import json
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True)
    ap.add_argument("--template", required=True)
    ap.add_argument("--output", required=True)
    ap.add_argument("--background", default="none", choices=["memory", "none"])
    a = ap.parse_args()
    print(json.dumps(generate(a.input, a.template, a.output, background=a.background),
                     ensure_ascii=False, indent=2))
# layout: reframe_chart() centres the table horizontally + starts it below the 부제
