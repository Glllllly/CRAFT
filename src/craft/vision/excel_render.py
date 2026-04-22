
"""
Excel Screenshot -> (optional) Grid Overlay -> Pixel Points -> Excel Cell Anchors

Point-only design (no bounding boxes):
- Model outputs 3 anchor points per region.
- You map pixel points -> Excel (row, col) via row/col pixel edges.
- Includes "snapping" to nearest interesting cell + merged-cell anchoring.
- `grid_alpha` controls grid overlay opacity (0-255).
- `grid_rgb` controls grid overlay color (RGB).
"""

import os
import time
import json
import bisect
import sys
from collections import deque, Counter
from time import perf_counter

import openpyxl
from PIL import Image, ImageDraw

WALL_STYLES = {"medium", "thick", "double"}
# Treat thin borders as "structure" too; otherwise snapping often misses tables.
STRUCT_BORDER = {"thin", "medium", "thick", "double", "hair"}
DEFAULT_DPI = 96


def export_sheet_png_via_excel_com(
    xlsx_path: str,
    out_png: str,
    sheet_name: str | None = None,
    dpi: int = DEFAULT_DPI,
    zoom: int = 100,
) -> tuple[int, int, int, int, list[float], list[float]]:
    """Windows only. Requires: Excel + pywin32 + desktop session."""
    try:
        import win32com.client as win32
    except Exception as e:
        raise RuntimeError("pywin32/win32com not available; cannot export via Excel COM.") from e
    from PIL import ImageGrab

    t0 = perf_counter()
    print("Step: render worksheet", flush=True)

    excel = None
    wb = None
    try:
        # DispatchEx creates a new instance (avoids attaching to a busy Excel).
        excel = win32.DispatchEx("Excel.Application")
        # NOTE: Some Excel versions return a blank CopyPicture when running fully invisible.
        # Keep the app minimized but visible to force rendering.
        excel.Visible = True
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.ScreenUpdating = True
        try:
            # 2 = xlMinimized
            excel.WindowState = 2
        except Exception:
            pass
        # 3 = msoAutomationSecurityForceDisable (disable macros for automation)
        try:
            excel.AutomationSecurity = 3
        except Exception:
            pass

        wb = excel.Workbooks.Open(
            os.path.abspath(xlsx_path),
            UpdateLinks=0,
            ReadOnly=True,
            AddToMru=False,
        )
        ws = wb.Worksheets(sheet_name) if sheet_name else wb.Worksheets(1)

        try:
            excel.ActiveWindow.Zoom = zoom
        except Exception:
            pass

        try:
            ws.Activate()
        except Exception:
            pass

        used = ws.UsedRange
        try:
            used.Select()
        except Exception:
            pass
        used.CopyPicture(Appearance=1, Format=2)  # 2 = bitmap

        # Clipboard is async; poll with timeout to avoid "hang forever".
        img = None
        deadline = perf_counter() + 10.0
        while perf_counter() < deadline:
            time.sleep(0.25)
            img = ImageGrab.grabclipboard()
            if img is not None:
                break
        if img is None:
            raise RuntimeError(
                "Clipboard grab failed or timed out. Ensure you're running in an interactive desktop session "
                "(not via service/ssh), and try closing any Excel instances then rerun."
            )

        img.save(out_png)

        min_row = int(used.Row)
        min_col = int(used.Column)
        max_row = min_row + int(used.Rows.Count) - 1
        max_col = min_col + int(used.Columns.Count) - 1

        # Compute exact row/col edges in the same coordinate system as the screenshot.
        origin = ws.Cells(min_row, min_col)
        origin_left = float(origin.Left)
        origin_top = float(origin.Top)
        pt_to_px = float(dpi) / 72.0

        # Column edges: use each column's Left, plus the last column's Right.
        x_edges_px: list[float] = []
        for c in range(min_col, max_col + 1):
            left_pt = float(ws.Cells(min_row, c).Left) - origin_left
            x_edges_px.append(left_pt * pt_to_px)
        last = ws.Cells(min_row, max_col)
        x_edges_px.append(float(last.Left + last.Width - origin_left) * pt_to_px)

        # Row edges: use each row's Top, plus the last row's Bottom.
        y_edges_px: list[float] = []
        for r in range(min_row, max_row + 1):
            top_pt = float(ws.Cells(r, min_col).Top) - origin_top
            y_edges_px.append(top_pt * pt_to_px)
        last = ws.Cells(max_row, min_col)
        y_edges_px.append(float(last.Top + last.Height - origin_top) * pt_to_px)

        return min_row, min_col, max_row, max_col, x_edges_px, y_edges_px
    finally:
        try:
            if wb is not None:
                wb.Close(False)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass


def get_used_bbox(ws):
    """
    Try to avoid pathological `calculate_dimension()` results when sheets have far-away formatting.

    Prefers `ws._cells` (only materialized cells) and falls back to `calculate_dimension()`.
    """
    cells = getattr(ws, "_cells", None)
    if cells:
        rows = [rc[0] for rc in cells.keys()]
        cols = [rc[1] for rc in cells.keys()]
        return min(rows), min(cols), max(rows), max(cols)

    dim = ws.calculate_dimension()  # e.g. "A1:U44"
    a, b = dim.split(":")
    r1, c1 = openpyxl.utils.coordinate_to_tuple(a)
    r2, c2 = openpyxl.utils.coordinate_to_tuple(b)
    return r1, c1, r2, c2


def col_width_to_px(width):
    if width is None:
        width = 8.43
    return int(width * 7 + 5)


def row_height_to_px(height, dpi=DEFAULT_DPI):
    if height is None:
        height = 15.0
    return int(height * dpi / 72)


def compute_edges(ws, min_row, min_col, max_row, max_col, dpi=DEFAULT_DPI):
    """Approximate edges from openpyxl widths/heights (fast, but may drift vs Excel rendering)."""
    x_edges = [0]
    for c in range(min_col, max_col + 1):
        letter = openpyxl.utils.get_column_letter(c)
        w = ws.column_dimensions[letter].width
        x_edges.append(x_edges[-1] + col_width_to_px(w))

    y_edges = [0]
    for r in range(min_row, max_row + 1):
        h = ws.row_dimensions[r].height
        y_edges.append(y_edges[-1] + row_height_to_px(h, dpi=dpi))

    return x_edges, y_edges


def compute_edges_via_excel_com(
    xlsx_path: str,
    sheet_name: str | None,
    min_row: int,
    min_col: int,
    max_row: int,
    max_col: int,
    dpi: int = DEFAULT_DPI,
    zoom: int = 100,
):
    """Exact edges from Excel COM (matches what Excel actually renders).

    Key fix vs the earlier version:
    - Don't just *sum widths/heights*.
      Excel's rendered grid can have tiny per-column/row deviations (and there are also
      left/top offsets). So we query each column/row's *Left/Top* position directly.

    Returns:
      x_edges, y_edges in *pixels* (float), relative to the top-left of (min_row, min_col).
    """
    try:
        import win32com.client as win32
    except Exception as e:
        raise RuntimeError("pywin32/win32com not available; cannot compute COM edges.") from e

    excel = None
    wb = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        try:
            excel.AutomationSecurity = 3
        except Exception:
            pass

        wb = excel.Workbooks.Open(
            os.path.abspath(xlsx_path),
            UpdateLinks=0,
            ReadOnly=True,
            AddToMru=False,
        )
        ws = wb.Worksheets(sheet_name) if sheet_name else wb.Worksheets(1)

        try:
            excel.ActiveWindow.Zoom = zoom
        except Exception:
            pass

        origin = ws.Cells(min_row, min_col)
        origin_left = float(origin.Left)
        origin_top = float(origin.Top)
        pt_to_px = float(dpi) / 72.0

        # Column edges: use each column's Left, plus the last column's Right.
        x_edges: list[float] = []
        for c in range(min_col, max_col + 1):
            x_edges.append((float(ws.Cells(min_row, c).Left) - origin_left) * pt_to_px)
        last = ws.Cells(min_row, max_col)
        x_edges.append((float(last.Left + last.Width) - origin_left) * pt_to_px)

        # Row edges: use each row's Top, plus the last row's Bottom.
        y_edges: list[float] = []
        for r in range(min_row, max_row + 1):
            y_edges.append((float(ws.Cells(r, min_col).Top) - origin_top) * pt_to_px)
        last = ws.Cells(max_row, min_col)
        y_edges.append((float(last.Top + last.Height) - origin_top) * pt_to_px)

        return x_edges, y_edges
    finally:
        try:
            if wb is not None:
                wb.Close(False)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass


def draw_grid_overlay(
    png_in: str,
    png_out: str,
    x_edges,
    y_edges,
    stride: int = 1,
    alpha: int = 0,
    rgb: tuple[int, int, int] = (0, 160, 0),
):
    img = Image.open(png_in).convert("RGBA")
    draw = ImageDraw.Draw(img)

    alpha = int(max(0, min(255, alpha)))
    r, g, b = (int(rgb[0]), int(rgb[1]), int(rgb[2]))
    r = max(0, min(255, r))
    g = max(0, min(255, g))
    b = max(0, min(255, b))
    line_color = (r, g, b, alpha)

    for i, x in enumerate(x_edges):
        if stride > 1 and i % stride != 0:
            continue
        xx = int(round(float(x)))
        draw.line([(xx, 0), (xx, img.height)], fill=line_color, width=1)

    for i, y in enumerate(y_edges):
        if stride > 1 and i % stride != 0:
            continue
        yy = int(round(float(y)))
        draw.line([(0, yy), (img.width, yy)], fill=line_color, width=1)

    img.save(png_out)


def point_to_excel_rc(x, y, x_edges, y_edges, min_row, min_col):
    """Pixel (x,y) -> Excel (row,col).

    Assumes:
      - x_edges/y_edges are aligned to `sheet.png` pixel space
      - min_row/min_col are the UsedRange top-left

    Note: This is the "reverse parse" you want.
    """
    c0 = bisect.bisect_right(x_edges, x) - 1
    r0 = bisect.bisect_right(y_edges, y) - 1
    c0 = max(0, min(c0, len(x_edges) - 2))
    r0 = max(0, min(r0, len(y_edges) - 2))
    return (min_row + r0, min_col + c0)


def excel_rc_to_a1(r, c):
    """(row,col) -> 'A1' style address."""
    return f"{openpyxl.utils.get_column_letter(int(c))}{int(r)}"


def build_merge_anchor_map(ws):
    m = {}
    for mr in ws.merged_cells.ranges:
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                m[(r, c)] = (mr.min_row, mr.min_col)
    return m


def has_struct_border(cell):
    for side in ["top", "bottom", "left", "right"]:
        s = getattr(cell.border, side)
        if s is not None and s.style is not None and str(s.style).lower() in STRUCT_BORDER:
            return True
    return False

def has_meaningful_fill(cell):
    fill = cell.fill
    if not fill:
        return False
    # 没有 pattern 就基本是默认/空
    if fill.patternType is None:
        return False

    fg = fill.fgColor
    if not fg or fg.type != "rgb" or not fg.rgb:
        return False

    rgb = str(fg.rgb).upper()
    # 这些通常是“默认透明/默认白/默认黑”，按你任务一般不算“结构信号”
    if rgb in {"00000000", "FFFFFFFF", "FF000000"}:
        return False
    return True

def is_interesting_cell(ws, r, c, merge_anchor_map):
    if (r, c) in merge_anchor_map:
        return True
    cell = ws.cell(r, c)
    if cell.value not in (None, ""):
        return True
    if has_struct_border(cell):
        return True
    if has_meaningful_fill(cell):
        return True

    return False


def snap_to_nearest_interesting(ws, r, c, is_interesting, radius=4, max_row=None, max_col=None):
    # If the initially mapped cell isn't interesting, search outward.
    # Increase radius a bit for robustness on sparse tables.
    if radius < 6:
        radius = 6
    q = deque([(r, c, 0)])
    seen = {(r, c)}
    while q:
        cr, cc, d = q.popleft()
        if is_interesting(ws, cr, cc):
            return cr, cc
        if d >= radius:
            continue
        for nr, nc in [
            (cr + 1, cc),
            (cr - 1, cc),
            (cr, cc + 1),
            (cr, cc - 1),
            (cr + 1, cc + 1),
            (cr + 1, cc - 1),
            (cr - 1, cc + 1),
            (cr - 1, cc - 1),
        ]:
            if (nr, nc) in seen:
                continue
            if nr < 1 or nc < 1:
                continue
            if max_row is not None and nr > max_row:
                continue
            if max_col is not None and nc > max_col:
                continue
            seen.add((nr, nc))
            q.append((nr, nc, d + 1))
    return r, c


def anchor_from_points(
    ws,
    points,
    x_edges,
    y_edges,
    min_row,
    min_col,
    merge_anchor_map,
    is_interesting,
    radius=4,
    max_row=None,
    max_col=None,
):
    mapped = []
    for x, y in points:
        r, c = point_to_excel_rc(x, y, x_edges, y_edges, min_row, min_col)
        r, c = snap_to_nearest_interesting(
            ws, r, c, is_interesting, radius=radius, max_row=max_row, max_col=max_col
        )
        if (r, c) in merge_anchor_map:
            r, c = merge_anchor_map[(r, c)]
        mapped.append((r, c))

    # Use all points to pick the best anchor among candidates (more stable than majority vote).
    candidates = list(dict.fromkeys(mapped))  # stable unique
    best = None
    best_score = float("inf")
    for rr, cc in candidates:
        cx, cy = excel_rc_to_pixel_center(rr, cc, x_edges, y_edges, min_row, min_col)
        score = 0.0
        for x, y in points:
            dx = float(cx) - float(x)
            dy = float(cy) - float(y)
            score += dx * dx + dy * dy
        if score < best_score:
            best_score = score
            best = (rr, cc)

    if best is None:
        (r, c), _ = Counter(mapped).most_common(1)[0]
        best = (r, c)
    return best, mapped


def excel_rc_to_pixel_center(r, c, x_edges, y_edges, min_row, min_col):
    rr = r - min_row
    cc = c - min_col
    rr = max(0, min(rr, len(y_edges) - 2))
    cc = max(0, min(cc, len(x_edges) - 2))
    x = (float(x_edges[cc]) + float(x_edges[cc + 1])) / 2.0
    y = (float(y_edges[rr]) + float(y_edges[rr + 1])) / 2.0
    return x, y


def excel_rc_to_pixel_box(r, c, x_edges, y_edges, min_row, min_col):
    rr = r - min_row
    cc = c - min_col
    rr = max(0, min(rr, len(y_edges) - 2))
    cc = max(0, min(cc, len(x_edges) - 2))
    x0 = float(x_edges[cc])
    x1 = float(x_edges[cc + 1])
    y0 = float(y_edges[rr])
    y1 = float(y_edges[rr + 1])
    return x0, y0, x1, y1


def draw_points_and_anchors(png_in, png_out, regions, anchors, x_edges, y_edges, min_row, min_col):
    img = Image.open(png_in).convert("RGBA")
    draw = ImageDraw.Draw(img)

    for reg in regions:
        for (x, y) in reg["anchor_points"]:
            draw.ellipse((x - 4, y - 4, x + 4, y + 4), outline=(255, 0, 0, 200), width=2)

    for anc in anchors:
        r, c = anc["anchor_cell"]
        x, y = excel_rc_to_pixel_center(r, c, x_edges, y_edges, min_row, min_col)
        x = int(round(x))
        y = int(round(y))
        draw.ellipse((x - 5, y - 5, x + 5, y + 5), outline=(0, 120, 255, 220), width=3)
        x0, y0, x1, y1 = excel_rc_to_pixel_box(r, c, x_edges, y_edges, min_row, min_col)
        draw.rectangle(
            (int(round(x0)), int(round(y0)), int(round(x1)), int(round(y1))),
            outline=(0, 120, 255, 180),
            width=1,
        )

    img.save(png_out)


def run_pipeline(
    xlsx_path: str,
    out_dir: str = "out",
    sheet_name: str | None = None,
    dpi: int = DEFAULT_DPI,
    grid_stride: int = 1,
    grid_alpha: int = 255,
    grid_rgb: tuple[int, int, int] = (0, 160, 0),
    points_filename: str = "points.json",
    prefer_excel_edges: bool = True,
):
    t0 = perf_counter()
    print("Step: render worksheet", flush=True)
    if not os.path.isabs(out_dir):
        out_dir = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), out_dir))
    os.makedirs(out_dir, exist_ok=True)

    sheet_png = os.path.join(out_dir, "sheet.png")
    sheet_grid_png = os.path.join(out_dir, "sheet_grid.png")
    anchors_json = os.path.join(out_dir, "anchors.json")
    debug_png = os.path.join(out_dir, "sheet_anchors.png")
    bounds_json = os.path.join(out_dir, "sheet_bounds.json")
    edges_json = os.path.join(out_dir, "sheet_edges.json")

    min_row = min_col = max_row = max_col = None
    x_edges = y_edges = None
    if not os.path.exists(sheet_png):
        min_row, min_col, max_row, max_col, x_edges, y_edges = export_sheet_png_via_excel_com(
            xlsx_path, sheet_png, sheet_name=sheet_name, dpi=dpi
        )
        with open(bounds_json, "w", encoding="utf-8") as f:
            json.dump(
                {"min_row": min_row, "min_col": min_col, "max_row": max_row, "max_col": max_col},
                f,
                ensure_ascii=False,
                indent=2,
            )
        with open(edges_json, "w", encoding="utf-8") as f:
            json.dump(
                {"dpi": dpi, "x_edges": x_edges, "y_edges": y_edges},
                f,
                ensure_ascii=False,
            )

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active

    if os.path.exists(bounds_json):
        with open(bounds_json, "r", encoding="utf-8") as f:
            b = json.load(f)
        min_row, min_col, max_row, max_col = b["min_row"], b["min_col"], b["max_row"], b["max_col"]
    else:
        min_row, min_col, max_row, max_col = get_used_bbox(ws)

    # Prefer exact Excel-rendered geometry when possible (fixes drifting / skewed grid overlays).
    if x_edges is None or y_edges is None:
        if os.path.exists(edges_json):
            with open(edges_json, "r", encoding="utf-8") as f:
                e = json.load(f)
            x_edges, y_edges = e["x_edges"], e["y_edges"]
        elif prefer_excel_edges:
            try:
                x_edges, y_edges = compute_edges_via_excel_com(
                    xlsx_path=xlsx_path,
                    sheet_name=sheet_name,
                    min_row=min_row,
                    min_col=min_col,
                    max_row=max_row,
                    max_col=max_col,
                    dpi=dpi,
                )
                with open(edges_json, "w", encoding="utf-8") as f:
                    json.dump({"dpi": dpi, "x_edges": x_edges, "y_edges": y_edges}, f, ensure_ascii=False)
            except Exception as _e:
                x_edges, y_edges = compute_edges(ws, min_row, min_col, max_row, max_col, dpi=dpi)
        else:
            x_edges, y_edges = compute_edges(ws, min_row, min_col, max_row, max_col, dpi=dpi)

    # --- IMPORTANT: align edges to the exported PNG pixel space ---
# CopyPicture output size depends on zoom / Windows display scaling.
# So we *always* rescale edges to exactly match the exported image size.
    img = Image.open(sheet_png)
    if x_edges[-1] > 0 and y_edges[-1] > 0:
        sx = img.width / float(x_edges[-1])
        sy = img.height / float(y_edges[-1])
        # keep as float; only round at draw-time (reduces accumulated drift)
        x_edges = [x * sx for x in x_edges]
        y_edges = [y * sy for y in y_edges]
        x_edges[-1] = float(img.width)
        y_edges[-1] = float(img.height)


    draw_grid_overlay(
        sheet_png,
        sheet_grid_png,
        x_edges,
        y_edges,
        stride=grid_stride,
        alpha=grid_alpha,
        rgb=grid_rgb,
    )
    points_path = os.path.join(out_dir, points_filename)
    if not os.path.exists(points_path):
        print("Step: worksheet render complete", flush=True)
        return

    with open(points_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    print("Step: map points to cells", flush=True)

    if isinstance(data, dict):
        regions = data.get("regions", [])
    elif isinstance(data, list):
        regions = data
    else:
        raise TypeError(f"Unsupported points.json format: {type(data)} (expected object or array)")

    merge_anchor_map = build_merge_anchor_map(ws)
    interesting_cache: dict[tuple[int, int], bool] = {}

    def is_interesting_cached(ws_, r, c):
        key = (r, c)
        hit = interesting_cache.get(key)
        if hit is not None:
            return hit
        val = is_interesting_cell(ws_, r, c, merge_anchor_map)
        interesting_cache[key] = val
        return val

    anchors = []
    for idx, reg in enumerate(regions):
        anchor, mapped = anchor_from_points(
            ws,
            reg["anchor_points"],
            x_edges, y_edges,
            min_row, min_col,
            merge_anchor_map=merge_anchor_map,
            is_interesting=is_interesting_cached,
            radius=4,
            max_row=max_row,
            max_col=max_col,
        )
        anchors.append({
            "region_id": idx,
            "type": reg.get("type", "unknown"),
            "anchor_cell": list(anchor),
            "mapped_cells_debug": [list(x) for x in mapped],
            "comment": reg.get("comment", "")
        })
    with open(anchors_json, "w", encoding="utf-8") as f:
        json.dump({"anchors": anchors}, f, ensure_ascii=False, indent=2)

    # Draw on the grid-overlay image so the final debug render preserves the visible sheet grid.
    debug_base_png = sheet_grid_png if os.path.exists(sheet_grid_png) else sheet_png
    draw_points_and_anchors(debug_base_png, debug_png, regions, anchors, x_edges, y_edges, min_row, min_col)

    print("Step: worksheet mapping complete", flush=True)


if __name__ == "__main__":
    here = os.path.dirname(os.path.abspath(__file__))
    run_pipeline(
        xlsx_path=os.path.join(here, "input.xlsx"),
        out_dir=os.path.join(here, "excel_out"),
        sheet_name=None,
        dpi=96,
        grid_stride=1,
        grid_alpha=255,
        points_filename="points.json",
    )
