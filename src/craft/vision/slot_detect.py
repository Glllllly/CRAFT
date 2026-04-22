# full_pipeline_seg2cell_qwen.py
# pip install torch torchvision transformers timm numpy pillow opencv-python openpyxl openai

import os
import json
import base64
import argparse
import re
import sys
from pathlib import Path
from typing import List, Dict, Tuple, Any

import cv2
import numpy as np
import torch
import torch.nn.functional as F
from PIL import Image
import openpyxl
from openai import OpenAI

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)


def resolve_default_segformer_ckpt() -> str:
    root = Path(ROOT_DIR)
    candidates = (
        root / "further_further_train" / "split_runs_recall_boost_new" / "best.pt",
        root / "further_further_train" / "split_runs" / "best.pt",
        root / "further_further_train" / "ckpt" / "best-097.pt",
        Path(__file__).resolve().parent / "best_segformer.pt",
        root.parent / "ckpt_cf" / "best_segformer_multitask.pt",
        root.parent / "ckpt_cf" / "new" / "best_segformer_multitask.pt",
        root.parent / "ckpt_cf" / "new" / "best_segformer.pt",
        Path.cwd() / "best.pt",
        Path.cwd() / "best_segformer_multitask.pt",
    )
    for candidate in candidates:
        if candidate.is_file():
            return str(candidate)
    return str(candidates[0])


def _connected_component_bboxes(mask_bin: np.ndarray, min_area: int = 16) -> List[Tuple[int, int, int, int]]:
    height, width = mask_bin.shape
    visited = np.zeros((height, width), dtype=bool)
    boxes: List[Tuple[int, int, int, int]] = []
    for y in range(height):
        for x in range(width):
            if mask_bin[y, x] == 0 or visited[y, x]:
                continue
            stack = [(y, x)]
            visited[y, x] = True
            min_y = max_y = y
            min_x = max_x = x
            area = 0
            while stack:
                cy, cx = stack.pop()
                area += 1
                min_y = min(min_y, cy)
                max_y = max(max_y, cy)
                min_x = min(min_x, cx)
                max_x = max(max_x, cx)
                for ny, nx in ((cy - 1, cx), (cy + 1, cx), (cy, cx - 1), (cy, cx + 1)):
                    if ny < 0 or ny >= height or nx < 0 or nx >= width:
                        continue
                    if visited[ny, nx] or mask_bin[ny, nx] == 0:
                        continue
                    visited[ny, nx] = True
                    stack.append((ny, nx))
            if area >= min_area:
                boxes.append((min_y, min_x, max_y, max_x))
    return boxes


def _pick_edge_index(scores: np.ndarray, start: int, end: int, fallback: int) -> int:
    start = max(0, min(int(start), scores.shape[0] - 1))
    end = max(start + 1, min(int(end), scores.shape[0]))
    window = scores[start:end]
    if window.size == 0:
        return int(fallback)
    return int(start + int(window.argmax()))


def _rectify_mask_from_probs_local(
    final_prob: np.ndarray,
    horizontal_prob: np.ndarray,
    vertical_prob: np.ndarray,
    final_threshold: float,
    horizontal_threshold: float,
    vertical_threshold: float,
    rectify_alpha: float,
    rectify_max_shift: int,
    rectify_min_edge_score_ratio: float,
    rectify_pad: int,
    rectify_min_area: int,
) -> np.ndarray:
    final_mask = (final_prob >= float(final_threshold)).astype(np.uint8)
    horizontal_support = np.maximum(horizontal_prob - float(horizontal_threshold), 0.0)
    vertical_support = np.maximum(vertical_prob - float(vertical_threshold), 0.0)
    rectified = np.zeros_like(final_mask, dtype=np.uint8)
    height, width = final_mask.shape
    for y0, x0, y1, x1 in _connected_component_bboxes(final_mask, min_area=max(1, int(rectify_min_area))):
        sy0 = max(0, y0 - int(rectify_pad))
        sx0 = max(0, x0 - int(rectify_pad))
        sy1 = min(height - 1, y1 + int(rectify_pad))
        sx1 = min(width - 1, x1 + int(rectify_pad))
        row_scores = horizontal_support[sy0 : sy1 + 1, sx0 : sx1 + 1].sum(axis=1)
        col_scores = vertical_support[sy0 : sy1 + 1, sx0 : sx1 + 1].sum(axis=0)
        cy = (y0 + y1) // 2
        cx = (x0 + x1) // 2
        cy_rel = cy - sy0
        cx_rel = cx - sx0
        top = _pick_edge_index(row_scores, 0, cy_rel + 1, y0 - sy0) + sy0
        bottom = _pick_edge_index(row_scores, cy_rel, row_scores.shape[0], y1 - sy0) + sy0
        left = _pick_edge_index(col_scores, 0, cx_rel + 1, x0 - sx0) + sx0
        right = _pick_edge_index(col_scores, cx_rel, col_scores.shape[0], x1 - sx0) + sx0
        row_peak = float(row_scores.max()) if row_scores.size else 0.0
        col_peak = float(col_scores.max()) if col_scores.size else 0.0
        row_mean = float(row_scores.mean()) if row_scores.size else 0.0
        col_mean = float(col_scores.mean()) if col_scores.size else 0.0
        row_ok = row_peak > 0.0 and row_peak >= float(rectify_min_edge_score_ratio) * max(1e-6, row_mean)
        col_ok = col_peak > 0.0 and col_peak >= float(rectify_min_edge_score_ratio) * max(1e-6, col_mean)
        if row_ok:
            top = int(round((1.0 - float(rectify_alpha)) * y0 + float(rectify_alpha) * top))
            bottom = int(round((1.0 - float(rectify_alpha)) * y1 + float(rectify_alpha) * bottom))
            top = max(y0 - int(rectify_max_shift), min(y0 + int(rectify_max_shift), top))
            bottom = max(y1 - int(rectify_max_shift), min(y1 + int(rectify_max_shift), bottom))
        else:
            top, bottom = y0, y1
        if col_ok:
            left = int(round((1.0 - float(rectify_alpha)) * x0 + float(rectify_alpha) * left))
            right = int(round((1.0 - float(rectify_alpha)) * x1 + float(rectify_alpha) * right))
            left = max(x0 - int(rectify_max_shift), min(x0 + int(rectify_max_shift), left))
            right = max(x1 - int(rectify_max_shift), min(x1 + int(rectify_max_shift), right))
        else:
            left, right = x0, x1
        if bottom < top:
            top, bottom = y0, y1
        if right < left:
            left, right = x0, x1
        rectified[top : bottom + 1, left : right + 1] = 1
    return rectified

try:
    from seg_model import (
        DEFAULT_PRETRAINED_MODEL,
        DEFAULT_FINAL_THRESHOLD,
        DEFAULT_HORIZONTAL_THRESHOLD,
        DEFAULT_VERTICAL_THRESHOLD,
        DEFAULT_RECTIFY_ALPHA,
        DEFAULT_RECTIFY_MAX_SHIFT,
        DEFAULT_RECTIFY_MIN_EDGE_SCORE_RATIO,
        DEFAULT_RECTIFY_PAD,
        DEFAULT_RECTIFY_MIN_AREA,
        RectangleAwareSegformer,
        load_flexible_checkpoint,
        predict_mask as predict_rectified_mask,
    )
except ImportError:
    import seg_model as _ras

    # Backward-compatible fallback for environments where qwen_small.py has been
    # updated but rectangle_aware_segformer.py is still an older revision.
    DEFAULT_PRETRAINED_MODEL = getattr(
        _ras, "DEFAULT_PRETRAINED_MODEL", "nvidia/segformer-b1-finetuned-ade-512-512"
    )
    DEFAULT_FINAL_THRESHOLD = getattr(_ras, "DEFAULT_FINAL_THRESHOLD", 0.45)
    DEFAULT_HORIZONTAL_THRESHOLD = getattr(_ras, "DEFAULT_HORIZONTAL_THRESHOLD", 0.20)
    DEFAULT_VERTICAL_THRESHOLD = getattr(_ras, "DEFAULT_VERTICAL_THRESHOLD", 0.10)
    DEFAULT_RECTIFY_ALPHA = getattr(_ras, "DEFAULT_RECTIFY_ALPHA", 0.50)
    DEFAULT_RECTIFY_MAX_SHIFT = getattr(_ras, "DEFAULT_RECTIFY_MAX_SHIFT", 16)
    DEFAULT_RECTIFY_MIN_EDGE_SCORE_RATIO = getattr(_ras, "DEFAULT_RECTIFY_MIN_EDGE_SCORE_RATIO", 0.50)
    DEFAULT_RECTIFY_PAD = getattr(_ras, "DEFAULT_RECTIFY_PAD", 4)
    DEFAULT_RECTIFY_MIN_AREA = getattr(_ras, "DEFAULT_RECTIFY_MIN_AREA", 16)
    RectangleAwareSegformer = _ras.RectangleAwareSegformer
    load_flexible_checkpoint = _ras.load_flexible_checkpoint
    if hasattr(_ras, "predict_mask"):
        predict_rectified_mask = _ras.predict_mask
    else:
        def predict_rectified_mask(
            model,
            image,
            device,
            size=512,
            final_threshold=DEFAULT_FINAL_THRESHOLD,
            horizontal_threshold=DEFAULT_HORIZONTAL_THRESHOLD,
            vertical_threshold=DEFAULT_VERTICAL_THRESHOLD,
            rectify_alpha=DEFAULT_RECTIFY_ALPHA,
            rectify_max_shift=DEFAULT_RECTIFY_MAX_SHIFT,
            rectify_min_edge_score_ratio=DEFAULT_RECTIFY_MIN_EDGE_SCORE_RATIO,
            rectify_pad=DEFAULT_RECTIFY_PAD,
            rectify_min_area=DEFAULT_RECTIFY_MIN_AREA,
        ):
            img = image.convert("RGB").resize((size, size), resample=Image.BILINEAR)
            arr = np.asarray(img, dtype=np.float32) / 255.0
            pixel_values = torch.from_numpy(arr).permute(2, 0, 1).unsqueeze(0).to(device)
            outputs = model(pixel_values=pixel_values)
            target_size = image.size[::-1]
            final_prob = torch.sigmoid(
                F.interpolate(outputs["final_logits"], size=target_size, mode="bilinear", align_corners=False)
            )[0, 0].detach().cpu().numpy()
            horizontal_prob = torch.sigmoid(
                F.interpolate(outputs["horizontal_logits"], size=target_size, mode="bilinear", align_corners=False)
            )[0, 0].detach().cpu().numpy()
            vertical_prob = torch.sigmoid(
                F.interpolate(outputs["vertical_logits"], size=target_size, mode="bilinear", align_corners=False)
            )[0, 0].detach().cpu().numpy()
            if hasattr(_ras, "rectify_mask_from_probs"):
                return _ras.rectify_mask_from_probs(
                    final_prob=final_prob,
                    horizontal_prob=horizontal_prob,
                    vertical_prob=vertical_prob,
                    final_threshold=final_threshold,
                    horizontal_threshold=horizontal_threshold,
                    vertical_threshold=vertical_threshold,
                    rectify_alpha=rectify_alpha,
                    rectify_max_shift=rectify_max_shift,
                    rectify_min_edge_score_ratio=rectify_min_edge_score_ratio,
                    rectify_pad=rectify_pad,
                    rectify_min_area=rectify_min_area,
                )
            return _rectify_mask_from_probs_local(
                final_prob=final_prob,
                horizontal_prob=horizontal_prob,
                vertical_prob=vertical_prob,
                final_threshold=final_threshold,
                horizontal_threshold=horizontal_threshold,
                vertical_threshold=vertical_threshold,
                rectify_alpha=rectify_alpha,
                rectify_max_shift=rectify_max_shift,
                rectify_min_edge_score_ratio=rectify_min_edge_score_ratio,
                rectify_pad=rectify_pad,
                rectify_min_area=rectify_min_area,
            )

# ---- your reverse-parse helpers ----
# assumes excel_render.py is in the same folder or PYTHONPATH
from excel_render import (
    compute_edges,                 # openpyxl-based edges
    compute_edges_via_excel_com,    # optional (Windows+Excel)
    get_used_bbox,
    point_to_excel_rc,
    excel_rc_to_a1,
    build_merge_anchor_map,
    is_interesting_cell,
    anchor_from_points,
)

# -------------------------
# Utilities
# -------------------------
def encode_image_base64(image_path: str) -> Tuple[str, str]:
    img = Image.open(image_path)
    fmt = img.format.lower() if img.format else (image_path.split(".")[-1].lower())
    mime_map = {"jpg": "jpeg", "jpeg": "jpeg", "png": "png", "webp": "webp", "bmp": "bmp", "gif": "gif", "tif": "tiff", "tiff": "tiff"}
    mime = mime_map.get(fmt, "jpeg")
    with open(image_path, "rb") as f:
        data = f.read()
    return base64.b64encode(data).decode("utf-8"), mime

def make_messages(image_path: str, text: str) -> List[Dict[str, Any]]:
    b64, mime = encode_image_base64(image_path)
    return [{
        "role": "user",
        "content": [
            {"type": "image_url", "image_url": {"url": f"data:image/{mime};base64,{b64}"}},
            {"type": "text", "text": text}
        ]
    }]

def cluster_edges_to_image(x_edges: List[float], y_edges: List[float], img_w: int, img_h: int) -> Tuple[List[float], List[float]]:
    """
    Your pipeline scales edges to exported PNG pixel space by sx=img.width/x_edges[-1], sy=img.height/y_edges[-1].
    We'll do the same scaling here (works when edges are computed in Excel points->px space).
    """
    if x_edges[-1] > 0 and y_edges[-1] > 0:
        sx = img_w / float(x_edges[-1])
        sy = img_h / float(y_edges[-1])
        x_edges = [float(x) * sx for x in x_edges]
        y_edges = [float(y) * sy for y in y_edges]
        x_edges[-1] = float(img_w)
        y_edges[-1] = float(img_h)
    return x_edges, y_edges


def component_anchor_points(labels: np.ndarray, comp_id: int, stats_row: np.ndarray, k: int = 9) -> List[Tuple[int,int]]:
    """
    Pick multiple points inside the connected component:
    - centroid approx (bbox center)
    - plus a few random pixels from the component mask
    """
    x, y, w, h, area = stats_row.tolist()
    pts: List[Tuple[int,int]] = []
    # bbox-based stable points
    pts.append((int(x + 0.5*w), int(y + 0.5*h)))
    pts.append((int(x + 0.35*w), int(y + 0.35*h)))
    pts.append((int(x + 0.65*w), int(y + 0.35*h)))
    pts.append((int(x + 0.35*w), int(y + 0.65*h)))
    pts.append((int(x + 0.65*w), int(y + 0.65*h)))

    # sample some true component pixels (helps when bbox covers multiple cells)
    ys, xs = np.where(labels == comp_id)
    if len(xs) > 0:
        idx = np.linspace(0, len(xs)-1, num=min(k, len(xs)), dtype=int)
        for i in idx:
            pts.append((int(xs[i]), int(ys[i])))

    # de-dup
    pts2 = []
    seen = set()
    for p in pts:
        if p not in seen:
            pts2.append(p)
            seen.add(p)
    return pts2

def rc_span_to_a1(r0: int, c0: int, r1: int, c1: int) -> str:
    a1_0 = excel_rc_to_a1(r0, c0)
    a1_1 = excel_rc_to_a1(r1, c1)
    return a1_0 if (r0 == r1 and c0 == c1) else f"{a1_0}:{a1_1}"

# -------------------------
# SegFormer inference -> mask in original image space
# -------------------------
@torch.no_grad()
def seg_infer_mask(
    seg_model: RectangleAwareSegformer,
    img_path: str,
    device: str,
    img_size: int = 512,
    prob_thresh: float = -1.0,
    horizontal_threshold: float = DEFAULT_HORIZONTAL_THRESHOLD,
    vertical_threshold: float = DEFAULT_VERTICAL_THRESHOLD,
    rectify_alpha: float = DEFAULT_RECTIFY_ALPHA,
    rectify_max_shift: int = DEFAULT_RECTIFY_MAX_SHIFT,
    rectify_min_edge_score_ratio: float = DEFAULT_RECTIFY_MIN_EDGE_SCORE_RATIO,
    rectify_pad: int = DEFAULT_RECTIFY_PAD,
    rectify_min_area: int = DEFAULT_RECTIFY_MIN_AREA,
) -> np.ndarray:
    img_pil = Image.open(img_path).convert("RGB")
    final_threshold = (
        float(prob_thresh)
        if prob_thresh is not None and float(prob_thresh) >= 0.0
        else DEFAULT_FINAL_THRESHOLD
    )
    return predict_rectified_mask(
        seg_model,
        img_pil,
        device=device,
        size=img_size,
        final_threshold=final_threshold,
        horizontal_threshold=horizontal_threshold,
        vertical_threshold=vertical_threshold,
        rectify_alpha=rectify_alpha,
        rectify_max_shift=rectify_max_shift,
        rectify_min_edge_score_ratio=rectify_min_edge_score_ratio,
        rectify_pad=rectify_pad,
        rectify_min_area=rectify_min_area,
    )

# -------------------------
# mask -> regions -> excel cell alignment
# -------------------------

def filter_slots_by_area(slots, abs_min=200, rel_ratio=0.3):
    if not slots:
        return slots
    areas = np.array([s.get("area_px", 0) for s in slots], dtype=np.float32)
    med = float(np.median(areas))
    thr = max(float(abs_min), float(rel_ratio) * med)
    kept = [s for s in slots if float(s.get("area_px", 0)) >= thr]
    return kept

def mask_to_excel_slots(
    mask01_hw: np.ndarray,
    ws,
    x_edges: List[float],
    y_edges: List[float],
    min_row: int,
    min_col: int,
    max_row: int,
    max_col: int,
    min_area: int = 200,
    min_fill_frac: float = 0.1,
    split_by_cells: bool = True,
) -> List[Dict[str, Any]]:
    """
    Returns list of slots with:
      - slot_id
      - pred_bbox_px (xyxy)
      - anchor_points_px
      - value_anchor_cell (row,col) + value_anchor_a1
      - value_range_a1 (span of mapped points, rough)
    """
    merge_anchor_map = build_merge_anchor_map(ws)

    def _is_interesting(_ws, r, c):
        return is_interesting_cell(_ws, r, c, merge_anchor_map)

    slots: List[Dict[str, Any]] = []
    sid = 0
    if split_by_cells:
        H, W = mask01_hw.shape[:2]
        for r in range(min_row, max_row + 1):
            rr = r - min_row
            y0 = int(round(float(y_edges[rr])))
            y1 = int(round(float(y_edges[rr + 1])))
            y0 = max(0, min(y0, H))
            y1 = max(0, min(y1, H))
            if y1 <= y0:
                continue
            for c in range(min_col, max_col + 1):
                cc = c - min_col
                x0 = int(round(float(x_edges[cc])))
                x1 = int(round(float(x_edges[cc + 1])))
                x0 = max(0, min(x0, W))
                x1 = max(0, min(x1, W))
                if x1 <= x0:
                    continue

                cell_mask = mask01_hw[y0:y1, x0:x1]
                area = int(np.count_nonzero(cell_mask))
                cell_area = int((y1 - y0) * (x1 - x0))
                if area < min_area:
                    continue
                if cell_area <= 0 or (area / float(cell_area)) < min_fill_frac:
                    continue

                br, bc = r, c
                best_a1 = excel_rc_to_a1(br, bc)
                span_a1 = rc_span_to_a1(br, bc, br, bc)
                pts = [
                    (int((x0 + x1) / 2), int((y0 + y1) / 2)),
                    (int(x0 + 0.25 * (x1 - x0)), int(y0 + 0.25 * (y1 - y0))),
                    (int(x0 + 0.75 * (x1 - x0)), int(y0 + 0.25 * (y1 - y0))),
                    (int(x0 + 0.25 * (x1 - x0)), int(y0 + 0.75 * (y1 - y0))),
                    (int(x0 + 0.75 * (x1 - x0)), int(y0 + 0.75 * (y1 - y0))),
                ]

                slots.append({
                    "slot_id": sid,
                    "pred_bbox_px": [int(x0), int(y0), int(x1), int(y1)],
                    "area_px": int(area),
                    "anchor_points_px": [[int(px), int(py)] for (px, py) in pts],
                    "mapped_cells_rc": [[int(br), int(bc)]],
                    "value_anchor_cell_rc": [int(br), int(bc)],
                    "value_anchor_a1": best_a1,
                    "value_range_a1": span_a1,
                })
                sid += 1
    else:
        m = (mask01_hw.astype(np.uint8) * 255)

        # light close to connect broken strokes (optional)
        k = cv2.getStructuringElement(cv2.MORPH_RECT, (1,1))
        # m = cv2.morphologyEx(m, cv2.MORPH_CLOSE, k, iterations=1)

        num, labels, stats, _ = cv2.connectedComponentsWithStats(m, connectivity=8)

        for comp_id in range(1, num):
            x, y, w, h, area = stats[comp_id].tolist()
            if area < min_area:
                continue

            pts = component_anchor_points(labels, comp_id, stats[comp_id], k=9)

            best, mapped = anchor_from_points(
                ws=ws,
                points=pts,
                x_edges=x_edges,
                y_edges=y_edges,
                min_row=min_row,
                min_col=min_col,
                merge_anchor_map=merge_anchor_map,
                is_interesting=_is_interesting,
                radius=6,
                max_row=max_row,
                max_col=max_col,
            )

            # best anchor (single cell)
            br, bc = best
            best_a1 = excel_rc_to_a1(br, bc)

            # mapped span (rough range)
            rs = [r for (r,c) in mapped]
            cs = [c for (r,c) in mapped]
            r0, r1 = int(min(rs)), int(max(rs))
            c0, c1 = int(min(cs)), int(max(cs))
            span_a1 = rc_span_to_a1(r0, c0, r1, c1)

            slots.append({
                "slot_id": sid,
                "pred_bbox_px": [int(x), int(y), int(x+w), int(y+h)],
                "area_px": int(area),
                "anchor_points_px": [[int(px), int(py)] for (px,py) in pts],
                "mapped_cells_rc": [[int(r), int(c)] for (r,c) in mapped],
                "value_anchor_cell_rc": [int(br), int(bc)],
                "value_anchor_a1": best_a1,
                "value_range_a1": span_a1,
            })
            sid += 1
    
    # stable order: top-to-bottom, left-to-right by bbox
    slots = filter_slots_by_area(slots, abs_min=min_area, rel_ratio=0.3)

    slots.sort(key=lambda s: (s["pred_bbox_px"][1], s["pred_bbox_px"][0]))
    # reassign slot_id after sort (so overlay id matches list order)
    for i, s in enumerate(slots):
        s["slot_id"] = i
    return slots

def draw_overlay(
    img_path: str,
    mask01_hw: np.ndarray,
    slots: List[Dict[str,Any]],
    out_path: str,
    label_mode: str = "id",
    min_label_box_area: int = 300,
):
    img = cv2.cvtColor(cv2.imread(img_path), cv2.COLOR_BGR2RGB)
    overlay = img.copy()

    # red tint for mask
    overlay[mask01_hw == 1] = (0.55*overlay[mask01_hw == 1] + 0.45*np.array([255,0,0])).astype(np.uint8)

    for s in slots:
        x1,y1,x2,y2 = s["pred_bbox_px"]
        sid = s["slot_id"]
        a1 = s.get("value_anchor_a1") or s.get("value_range_a1") or ""
        if label_mode == "none":
            continue
        if label_mode == "id":
            label = f"{sid}"
        elif label_mode == "a1":
            label = f"{a1}"
        else:
            label = f"{sid}:{a1}"
        if (x2 - x1) * (y2 - y1) < max(0, int(min_label_box_area)):
            continue
        (tw, th), baseline = cv2.getTextSize(label, cv2.FONT_HERSHEY_SIMPLEX, 0.55, 2)
        cx = int((x1 + x2) / 2)
        cy = int((y1 + y2) / 2)
        tx = int(cx - tw / 2)
        ty = int(cy + th / 2)
        tx = max(0, min(tx, overlay.shape[1] - tw))
        ty = max(th, min(ty, overlay.shape[0] - baseline))
        # Black outline under yellow text improves readability on dense red masks.
        cv2.putText(overlay, label, (tx, ty),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.55, (0,0,0), 4, cv2.LINE_AA)
        cv2.putText(overlay, label, (tx, ty),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.55, (255,255,0), 2, cv2.LINE_AA)

    cv2.imwrite(out_path, cv2.cvtColor(overlay, cv2.COLOR_RGB2BGR))

# -------------------------
# Qwen judging (OpenAI compatible)
# -------------------------
def build_qwen_prompt(slots_batch: List[Dict[str,Any]], img_w: int, img_h: int) -> str:
    lines = []
    for s in slots_batch:
        sid = s["slot_id"]
        x1,y1,x2,y2 = s["pred_bbox_px"]
        a1 = s["value_range_a1"]
        lines.append(f'{sid}: bbox=[{x1},{y1},{x2},{y2}] value_cell="{a1}"')
    slot_block = "\n".join(lines)

    return f"""
你会得到一张表格/表单图片。
图中的红色半透明底色、黄色数字，以及任何为了标记候选 slot 而额外添加的高亮/描边，都是系统后加的 overlay，不属于原始表单内容。
这些红黄标记不是 key、不是 label、不是 value，也不是表单自带文字；尤其不要把黄色数字当成 key 或字段文本。
你的任务：对每个 slot 判断它是否真的是可填写字段（is_valid_slot），并给出它对应的 key 标签（优先基于视觉结构：同一行左侧、或上方表头、或左侧行头）。
每个 slot 的 slot_id 和对齐到 Excel 的 value_cell(A1 范围) 以“需要处理的 slots”文本列表为准，不要把图片里的 overlay 当成表格文本去读。
注意：说明/备注等长段落一般不是 key，除非它直接标注该 slot。
输出必须是严格 JSON，不要输出任何额外文字。

坐标系：图像大小 {img_w}x{img_h}，bbox 使用 [x1,y1,x2,y2]。

需要处理的 slots：
{slot_block}

输出 JSON schema：
{{
  "pairs": [
    {{
      "slot_id": <int>,
      "value_cell": <string>,             // 直接复制输入的 value_cell
      "is_valid_slot": <bool>,
      "key": {{
        "text": <string or "">,
        "bbox": [x1,y1,x2,y2]            // 找不到用 [-1,-1,-1,-1]
      }},
      "confidence": <float 0..1>,
      "reason": <string>
    }}
  ]
}}
约束：
- 每个 slot_id 必须出现且只出现一次
- confidence 必须在 0~1
""".strip()

def extract_json_object(text: str) -> Dict[str, Any]:
    # simple robust extractor
    i = text.find("{")
    j = text.rfind("}")
    if i == -1 or j == -1 or j <= i:
        raise ValueError("No JSON object found in response.")
    return json.loads(text[i:j+1])

def _sanitize_json_text(text: str) -> str:
    text = text.strip()
    if text.startswith("```"):
        text = text.strip("`").strip()
    i = text.find("{")
    j = text.rfind("}")
    if i == -1 or j == -1 or j <= i:
        return text
    text = text[i:j+1]
    text = text.replace("“", "\"").replace("”", "\"").replace("’", "'").replace("‘", "'")
    text = re.sub(r",\s*([}\]])", r"\1", text)
    return text

def _save_qwen_raw(out_dir: str | None, batch_id: int, content: str) -> None:
    if not out_dir:
        return
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, f"qwen_raw_{batch_id:03d}.txt")
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)

def run_qwen_batches(
    client: OpenAI,
    model_name: str,
    overlay_img_path: str,
    slots: List[Dict[str,Any]],
    batch_size: int = 10,
    temperature: float = 0.0,
    out_dir: str | None = None,
) -> Dict[str, Any]:
    img = Image.open(overlay_img_path)
    W, H = img.size

    all_pairs = []
    for st in range(0, len(slots), batch_size):
        batch = slots[st:st+batch_size]
        prompt = build_qwen_prompt(batch, W, H)
        messages = make_messages(overlay_img_path, prompt)
        try:
            resp = client.chat.completions.create(
                model=model_name,
                messages=messages,
                temperature=temperature,
                response_format={"type": "json_object"},
            )
        except Exception:
            resp = client.chat.completions.create(
                model=model_name,
                messages=messages,
                temperature=temperature,
            )
        content = resp.choices[0].message.content
        _save_qwen_raw(out_dir, st // batch_size, content)
        try:
            data = extract_json_object(content)
        except Exception:
            try:
                sanitized = _sanitize_json_text(content)
                data = json.loads(sanitized)
            except Exception:
                repair_prompt = (
                    "Your previous response was not valid JSON. "
                    "Return ONLY a valid JSON object that matches the required schema. "
                    "No markdown, no extra text."
                )
                messages = make_messages(overlay_img_path, prompt + "\n\n" + repair_prompt)
                try:
                    resp = client.chat.completions.create(
                        model=model_name,
                        messages=messages,
                        temperature=temperature,
                        response_format={"type": "json_object"},
                    )
                except Exception:
                    resp = client.chat.completions.create(
                        model=model_name,
                        messages=messages,
                        temperature=temperature,
                    )
                content = resp.choices[0].message.content
                _save_qwen_raw(out_dir, (st // batch_size) + 1000, content)
                try:
                    data = extract_json_object(content)
                except Exception:
                    data = {"pairs": []}

        # normalize / ensure coverage
        got = {p.get("slot_id") for p in data.get("pairs", []) if isinstance(p, dict)}
        need = {s["slot_id"] for s in batch}
        exist = {p["slot_id"]: p for p in data.get("pairs", []) if isinstance(p, dict) and "slot_id" in p}

        for sid in sorted(list(need)):
            if sid in exist:
                all_pairs.append(exist[sid])
            else:
                # fallback
                value_cell = next(x["value_range_a1"] for x in batch if x["slot_id"] == sid)
                all_pairs.append({
                    "slot_id": sid,
                    "value_cell": value_cell,
                    "is_valid_slot": False,
                    "key": {"text": "", "bbox": [-1,-1,-1,-1]},
                    "confidence": 0.0,
                    "reason": "missing slot_id in model output"
                })

    all_pairs.sort(key=lambda x: x.get("slot_id", 1e9))
    return {"pairs": all_pairs}

# -------------------------
# Main
# -------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--image", required=True, help="input table image (png/jpg)")
    ap.add_argument("--xlsx", required=True, help="original excel file (.xlsx)")
    ap.add_argument("--sheet", default=None, help="sheet name (default: active)")
    ap.add_argument("--sheet_from_image", action="store_true", help="derive sheet name from image filename")
    ap.add_argument("--sheet_regex", default=r"_([0-9]{3})_heet_", help="regex with 3-digit sheet id in image name")
    ap.add_argument("--ckpt", default=resolve_default_segformer_ckpt(), help="SegFormer checkpoint path")
    ap.add_argument("--out_dir", default="infer_out_full", help="output directory")
    ap.add_argument("--seg_prob_thresh", type=float, default=-1.0, help=">=0 to use softmax prob threshold for seg mask (e.g. 0.3); <0 keeps argmax")
    ap.add_argument("--final_threshold", type=float, default=DEFAULT_FINAL_THRESHOLD)
    ap.add_argument("--horizontal_threshold", type=float, default=DEFAULT_HORIZONTAL_THRESHOLD)
    ap.add_argument("--vertical_threshold", type=float, default=DEFAULT_VERTICAL_THRESHOLD)
    ap.add_argument("--rectify_alpha", type=float, default=DEFAULT_RECTIFY_ALPHA)
    ap.add_argument("--rectify_max_shift", type=int, default=DEFAULT_RECTIFY_MAX_SHIFT)
    ap.add_argument("--rectify_min_edge_score_ratio", type=float, default=DEFAULT_RECTIFY_MIN_EDGE_SCORE_RATIO)
    ap.add_argument("--rectify_pad", type=int, default=DEFAULT_RECTIFY_PAD)
    ap.add_argument("--rectify_min_area", type=int, default=DEFAULT_RECTIFY_MIN_AREA)

    ap.add_argument("--dpi", type=int, default=96, help="dpi for openpyxl edge compute fallback")
    ap.add_argument("--prefer_excel_com_edges", action="store_true", help="try Excel COM edges first (Windows only)")
    ap.add_argument("--no_excel_com_edges", action="store_true", help="disable Excel COM edges and use openpyxl widths")
    ap.add_argument("--edges_json", default=None, help="path to sheet_edges.json with x_edges/y_edges")
    ap.add_argument("--bounds_json", default=None, help="path to sheet_bounds.json with min/max rows/cols")
    ap.add_argument("--skip_qwen", action="store_true", help="stop after slots/overlay generation")
    ap.add_argument("--overlay_label_mode", default="id", choices=["id", "a1", "id_a1", "none"], help="overlay label text mode")
    ap.add_argument("--overlay_min_label_box_area", type=int, default=300, help="skip labels for boxes smaller than this area (px^2)")

    # Qwen server options (OpenAI compatible)
    ap.add_argument("--qwen_base_url", default=os.environ.get("QWEN_BASE_URL", "http://192.188.188.79:4086/v1"))
    ap.add_argument("--qwen_api_key", default=os.environ.get("QWEN_API_KEY", "EMPTY"))
    ap.add_argument("--qwen_model", default=os.environ.get("QWEN_MODEL", "Qwen/Qwen3-VL-8B-Instruct-GGUF"))
    ap.add_argument("--qwen_batch", type=int, default=10)
    args = ap.parse_args()

    os.makedirs(args.out_dir, exist_ok=True)

    # load segformer
    device = "cuda" if torch.cuda.is_available() else "cpu"
    seg = RectangleAwareSegformer(
        pretrained_name=DEFAULT_PRETRAINED_MODEL,
        init_mode="pretrained",
    ).to(device)
    load_stats = load_flexible_checkpoint(seg, args.ckpt, map_location="cpu")
    print("Step: detect slots")
    seg.eval()

    # infer mask in original image space
    final_thr = args.final_threshold
    if args.seg_prob_thresh is not None and float(args.seg_prob_thresh) >= 0.0:
        final_thr = float(args.seg_prob_thresh)
    mask01_hw = seg_infer_mask(
        seg,
        args.image,
        device=device,
        img_size=512,
        prob_thresh=final_thr,
        horizontal_threshold=args.horizontal_threshold,
        vertical_threshold=args.vertical_threshold,
        rectify_alpha=args.rectify_alpha,
        rectify_max_shift=args.rectify_max_shift,
        rectify_min_edge_score_ratio=args.rectify_min_edge_score_ratio,
        rectify_pad=args.rectify_pad,
        rectify_min_area=args.rectify_min_area,
    )

    # load workbook + sheet
    wb = openpyxl.load_workbook(args.xlsx, data_only=True)
    sheet_name = args.sheet
    if sheet_name is None and args.sheet_from_image:
        m = re.search(args.sheet_regex, os.path.basename(args.image))
        if m:
            sheet_idx = int(m.group(1))
            sheet_name = f"Sheet{sheet_idx}"
    ws = wb[sheet_name] if sheet_name else wb.active

    # used range bounds
    min_row, min_col, max_row, max_col = get_used_bbox(ws)
    if args.bounds_json:
        with open(args.bounds_json, "r", encoding="utf-8") as f:
            b = json.load(f)
        min_row = int(b["min_row"])
        min_col = int(b["min_col"])
        max_row = int(b["max_row"])
        max_col = int(b["max_col"])

    # compute edges (same logic as pipeline.py)
    x_edges = y_edges = None
    if args.edges_json:
        with open(args.edges_json, "r", encoding="utf-8") as f:
            data = json.load(f)
        x_edges = data.get("x_edges")
        y_edges = data.get("y_edges")

    use_excel_com = args.prefer_excel_com_edges or not args.no_excel_com_edges
    if x_edges is None or y_edges is None:
        if use_excel_com:
            try:
                x_edges, y_edges = compute_edges_via_excel_com(
                    xlsx_path=args.xlsx,
                    sheet_name=sheet_name,
                    min_row=min_row, min_col=min_col,
                    max_row=max_row, max_col=max_col,
                    dpi=args.dpi,
                )
            except Exception:
                x_edges, y_edges = None, None

    if x_edges is None or y_edges is None:
        try:
            x_edges, y_edges = compute_edges(ws, min_row, min_col, max_row, max_col, dpi=args.dpi)
        except Exception:
            x_edges, y_edges = None, None

    # scale edges to match the input image pixel space
    img_bgr = cv2.imread(args.image)
    H, W = img_bgr.shape[:2]
    x_edges, y_edges = cluster_edges_to_image(x_edges, y_edges, W, H)

    # mask -> slots aligned to excel cells
    slots = mask_to_excel_slots(
        mask01_hw=mask01_hw,
        ws=ws,
        x_edges=x_edges,
        y_edges=y_edges,
        min_row=min_row,
        min_col=min_col,
        max_row=max_row,
        max_col=max_col,
        min_area=200,
    )

    # save mask
    mask_path = os.path.join(args.out_dir, "mask.png")
    cv2.imwrite(mask_path, (mask01_hw*255).astype(np.uint8))

    # save slots aligned
    slots_path = os.path.join(args.out_dir, "slots_aligned.json")
    with open(slots_path, "w", encoding="utf-8") as f:
        json.dump({
            "image": args.image,
            "xlsx": args.xlsx,
            "sheet": args.sheet or ws.title,
            "image_size": {"w": W, "h": H},
            "used_range": {"min_row": min_row, "min_col": min_col, "max_row": max_row, "max_col": max_col},
            "slots": slots
        }, f, ensure_ascii=False, indent=2)

    # overlay
    overlay_path = os.path.join(args.out_dir, "overlay.png")
    draw_overlay(
        args.image,
        mask01_hw,
        slots,
        overlay_path,
        label_mode=args.overlay_label_mode,
        min_label_box_area=args.overlay_min_label_box_area,
    )

    if args.skip_qwen:
        print("Step: slot detection complete")
        return

    # call Qwen
    qwen_client = OpenAI(api_key=args.qwen_api_key, base_url=args.qwen_base_url, timeout=3600)
    qwen_out = run_qwen_batches(
        client=qwen_client,
        model_name=args.qwen_model,
        overlay_img_path=overlay_path,
        slots=slots,
        batch_size=args.qwen_batch,
        temperature=0.0,
        out_dir=args.out_dir,
    )

    qwen_path = os.path.join(args.out_dir, "qwen_pairs.json")
    with open(qwen_path, "w", encoding="utf-8") as f:
        json.dump({
            "image": args.image,
            "overlay": overlay_path,
            "pairs": qwen_out["pairs"]
        }, f, ensure_ascii=False, indent=2)

    print("Step: slot detection complete")

if __name__ == "__main__":
    main()
