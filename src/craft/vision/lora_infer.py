import argparse
import json
import os
import re
from collections import defaultdict
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from PIL import Image, ImageDraw, ImageFont
import torch
from transformers import AutoModelForVision2Seq, AutoProcessor
from peft import PeftModel

try:
    from transformers import Qwen3VLForConditionalGeneration
except Exception:
    Qwen3VLForConditionalGeneration = None


def load_json(path: str) -> Any:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def extract_slots(slots_json: str) -> List[Dict[str, Any]]:
    data = load_json(slots_json)
    if isinstance(data, dict) and "slots" in data:
        return data["slots"]
    if isinstance(data, list):
        return data
    raise ValueError("slots_json format not recognized.")


def build_prompt(slots: List[Dict[str, Any]], img_w: int, img_h: int, prompt_context_text: str = "") -> str:
    lines = []
    for s in slots:
        sid = s.get("slot_id", 0)
        bbox = s.get("pred_bbox_px", [0, 0, 0, 0])
        value_cell = (
            s.get("value_range_a1")
            or s.get("value_anchor_a1")
            or s.get("value_cell")
            or ""
        )
        x1, y1, x2, y2 = bbox
        lines.append(f'{sid}: bbox=[{x1},{y1},{x2},{y2}] value_cell="{value_cell}"')
    slot_block = "\n".join(lines)

    base_prompt = (
        "You will receive a form image with candidate value slots.\n"
        f"Image size: {img_w}x{img_h}, bbox format [x1,y1,x2,y2].\n"
        "All blue annotations are synthetic overlays added by the system, including both the blue numbers and the blue box lines/fills.\n"
        "Those blue overlays are NOT part of the original form content, NOT field values, and NOT key text.\n"
        "Do not copy, read, or interpret any blue overlay as spreadsheet content.\n"
        "If there is no black box or underline around a location, it is an invalid slot and should usually be marked is_valid_slot=false.\n"
        "If there are no texts surrounding the slot, it is an invalid slot and should usually be marked is_valid_slot=false.\n"
        "For each slot, decide if it is a valid fillable field and provide the key text.\n"
        "Important: for every slot with is_valid_slot=true, you MUST output a non-empty key.text copied from nearby non-blue form text.\n"
        "If you cannot identify the key text for a slot, set is_valid_slot=false instead of leaving key.text empty or null.\n"
        "Do not omit the key field from the JSON.\n"
        "Extra fields:\n"
        "- is_duplicated: whether the key appears multiple times in the form.\n"
        "- number: if duplicated, the 1-based index among duplicates; otherwise 0.\n"
        "- is_merge_cell: whether the value cell is within a merged cell.\n"
        "Output must be strict JSON matching the schema below.\n\n"
        f"Slots:\n{slot_block}\n\n"
        "JSON schema:\n"
        "{\n"
        '  "pairs": [\n'
        "    {\n"
        '      "slot_id": <int>,\n'
        '      "value_cell": <string>,\n'
        '      "is_valid_slot": <bool>,\n'
        '      "is_duplicated": <bool>,\n'
        '      "number": <int>,\n'
        '      "is_merge_cell": <bool>,\n'
        '      "key": {"text": <string|null>, "bbox": [<int>, <int>, <int>, <int>]}\n'
        "    }\n"
        "  ]\n"
        "}"
    )
    if str(prompt_context_text or "").strip():
        base_prompt += (
            "\n\nAdditional reflect risk context:\n"
            f"{str(prompt_context_text).strip()}\n"
            "Use this context to avoid wrong key-to-slot matching in risky regions."
        )
    return base_prompt


def extract_json_object(text: str) -> Dict[str, Any]:
    i = text.find("{")
    j = text.rfind("}")
    if i == -1 or j == -1 or j <= i:
        raise ValueError("No JSON object found in response.")
    return json.loads(text[i : j + 1])


def sanitize_json_text(text: str) -> str:
    text = text.strip()
    if text.startswith("```"):
        text = text.strip("`").strip()
    i = text.find("{")
    j = text.rfind("}")
    if i == -1 or j == -1 or j <= i:
        return text
    text = text[i : j + 1]
    text = text.replace("“", "\"").replace("”", "\"").replace("’", "'").replace("‘", "'")
    text = re.sub(r",\s*([}\]])", r"\1", text)
    return text


def chunk_list(items: List[Dict[str, Any]], chunk_size: int) -> List[List[Dict[str, Any]]]:
    if chunk_size <= 0:
        return [items]
    return [items[i : i + chunk_size] for i in range(0, len(items), chunk_size)]


def _get_slot_bbox(slot: Dict[str, Any]) -> Tuple[int, int, int, int]:
    bbox = slot.get("pred_bbox_px", [0, 0, 0, 0])
    if not isinstance(bbox, (list, tuple)) or len(bbox) < 4:
        return (0, 0, 0, 0)
    try:
        x1, y1, x2, y2 = [int(round(float(bbox[i]))) for i in range(4)]
    except Exception:
        return (0, 0, 0, 0)
    return x1, y1, x2, y2


def _union_bbox(slots: List[Dict[str, Any]]) -> Tuple[int, int, int, int]:
    boxes = [_get_slot_bbox(s) for s in slots]
    xs1 = [b[0] for b in boxes]
    ys1 = [b[1] for b in boxes]
    xs2 = [b[2] for b in boxes]
    ys2 = [b[3] for b in boxes]
    return min(xs1), min(ys1), max(xs2), max(ys2)


def _clamp_crop(x1: int, y1: int, x2: int, y2: int, img_w: int, img_h: int) -> Tuple[int, int, int, int]:
    x1 = max(0, min(x1, img_w - 1))
    y1 = max(0, min(y1, img_h - 1))
    x2 = max(x1 + 1, min(x2, img_w))
    y2 = max(y1 + 1, min(y2, img_h))
    return x1, y1, x2, y2


def _build_crop_for_chunk(
    chunk: List[Dict[str, Any]],
    img_w: int,
    img_h: int,
    left_pad: int,
    top_pad: int,
    right_pad: int,
    bottom_pad: int,
) -> Tuple[int, int, int, int]:
    x1, y1, x2, y2 = _union_bbox(chunk)
    return _clamp_crop(
        x1 - max(0, left_pad),
        y1 - max(0, top_pad),
        x2 + max(0, right_pad),
        y2 + max(0, bottom_pad),
        img_w,
        img_h,
    )


def _translate_chunk_slots_for_crop(
    chunk: List[Dict[str, Any]], crop_xyxy: Tuple[int, int, int, int]
) -> List[Dict[str, Any]]:
    cx1, cy1, _, _ = crop_xyxy
    out = []
    for s in chunk:
        ss = dict(s)
        x1, y1, x2, y2 = _get_slot_bbox(s)
        ss["pred_bbox_px"] = [x1 - cx1, y1 - cy1, x2 - cx1, y2 - cy1]
        out.append(ss)
    return out


def _resize_image_and_slots_if_needed(
    image: Image.Image,
    chunk_slots: List[Dict[str, Any]],
    max_image_size: int,
) -> Tuple[Image.Image, List[Dict[str, Any]]]:
    if not max_image_size or int(max_image_size) <= 0:
        return image, chunk_slots
    w, h = image.size
    if max(w, h) <= int(max_image_size):
        return image, chunk_slots
    scale = float(max_image_size) / float(max(w, h))
    out_img = image.resize((max(1, int(round(w * scale))), max(1, int(round(h * scale)))), Image.BICUBIC)
    out_slots: List[Dict[str, Any]] = []
    for s in chunk_slots:
        ss = dict(s)
        x1, y1, x2, y2 = _get_slot_bbox(s)
        ss["pred_bbox_px"] = [
            int(round(x1 * scale)),
            int(round(y1 * scale)),
            int(round(x2 * scale)),
            int(round(y2 * scale)),
        ]
        out_slots.append(ss)
    return out_img, out_slots


def _load_debug_font(sz: int):
    for name in ("arial.ttf", "DejaVuSans-Bold.ttf", "DejaVuSans.ttf"):
        try:
            return ImageFont.truetype(name, sz)
        except Exception:
            continue
    try:
        return ImageFont.load_default()
    except Exception:
        return None


def _draw_slot_ids(image: Image.Image, slots: List[Dict[str, Any]]) -> Image.Image:
    vis = image.copy().convert("RGB")
    draw = ImageDraw.Draw(vis, "RGBA")
    for s in slots:
        sid = s.get("slot_id", "?")
        x1, y1, x2, y2 = _get_slot_bbox(s)
        draw.rectangle([x1, y1, x2, y2], outline=(0, 102, 255, 72), width=1, fill=(0, 102, 255, 16))
        box_h = max(1, y2 - y1)
        font = _load_debug_font(max(12, min(36, int(box_h * 0.6))))
        txt = str(sid)
        try:
            tb = draw.textbbox((0, 0), txt, font=font)
            tw, th = tb[2] - tb[0], tb[3] - tb[1]
        except Exception:
            tw, th = (len(txt) * 8, 14)
        tx = x1 + max(2, ((x2 - x1) - tw) // 2)
        ty = y1 + max(2, ((y2 - y1) - th) // 2)
        for dx, dy in ((1, 0), (-1, 0), (0, 1), (0, -1)):
            draw.text((tx + dx, ty + dy), txt, fill=(255, 255, 255, 220), font=font)
        draw.text((tx, ty), txt, fill=(0, 102, 255, 235), font=font)
    return vis


def _slot_value_cell(slot: Dict[str, Any]) -> str:
    return str(slot.get("value_range_a1") or slot.get("value_anchor_a1") or slot.get("value_cell") or "")


def _slot_is_merge(slot: Dict[str, Any]) -> bool:
    if "is_merge_cell" in slot:
        return bool(slot.get("is_merge_cell"))
    return ":" in _slot_value_cell(slot)


def _normalize_chunk_output(data: Any, chunk_slots: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    if not isinstance(data, dict):
        return []
    pairs = data.get("pairs")
    if not isinstance(pairs, list):
        return []

    slot_map: Dict[int, Dict[str, Any]] = {}
    for s in chunk_slots:
        try:
            slot_map[int(s.get("slot_id"))] = s
        except Exception:
            continue

    out: List[Dict[str, Any]] = []
    for p in pairs:
        if not isinstance(p, dict):
            continue
        sid_raw = p.get("slot_id")
        try:
            sid = int(sid_raw)
        except Exception:
            continue
        slot = slot_map.get(sid, {})

        key_obj = p.get("key")
        key_text = None
        if isinstance(key_obj, dict):
            key_text = key_obj.get("text")
        if key_text is None:
            key_text = p.get("key_text")
        if key_text is None and isinstance(p.get("key"), str):
            key_text = p.get("key")
        if isinstance(key_text, str):
            key_text = key_text.strip() or None

        is_valid_slot = p.get("is_valid_slot")
        if is_valid_slot is None:
            is_valid_slot = key_text is not None
        else:
            is_valid_slot = bool(is_valid_slot)
        if key_text is None:
            is_valid_slot = False

        value_cell = p.get("value_cell")
        if not value_cell:
            value_cell = _slot_value_cell(slot)

        number = p.get("number", slot.get("number", 0))
        try:
            number = int(number)
        except Exception:
            number = 0

        key_bbox = [-1, -1, -1, -1]
        if isinstance(key_obj, dict) and isinstance(key_obj.get("bbox"), list) and len(key_obj.get("bbox")) == 4:
            key_bbox = key_obj.get("bbox")

        reason = str(p.get("reason", "chunk_lora_infer"))
        if key_text is None:
            if reason:
                reason = f"{reason}; missing_key_text"
            else:
                reason = "missing_key_text"

        out.append(
            {
                "slot_id": sid,
                "value_cell": str(value_cell or ""),
                "is_valid_slot": bool(is_valid_slot),
                "is_duplicated": bool(p.get("is_duplicated", slot.get("is_duplicated", False))),
                "number": number,
                "is_merge_cell": bool(p.get("is_merge_cell", _slot_is_merge(slot))),
                "key": {"text": key_text, "bbox": key_bbox},
                "confidence": float(p.get("confidence", 1.0) or 1.0),
                "reason": reason,
            }
        )

    return out


def _json_default(obj: Any) -> Any:
    if isinstance(obj, Path):
        return str(obj)
    return str(obj)


def _write_debug_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=_json_default), encoding="utf-8")


def _save_chunk_debug(
    debug_dir: Path | None,
    chunk_idx: int,
    crop_xyxy: Optional[Tuple[int, int, int, int]],
    infer_slots: List[Dict[str, Any]],
    prompt: str,
    prompt_text: str,
    raw_text: str,
    chunk_data: Any,
    normalized_pairs: List[Dict[str, Any]],
    infer_image: Image.Image,
    infer_image_for_model: Image.Image,
) -> None:
    if debug_dir is None:
        return
    chunk_dir = debug_dir / f"chunk_{chunk_idx:03d}"
    chunk_dir.mkdir(parents=True, exist_ok=True)
    infer_image.save(chunk_dir / "crop.png")
    infer_image_for_model.save(chunk_dir / "model_input.png")
    (chunk_dir / "prompt.txt").write_text(prompt, encoding="utf-8")
    (chunk_dir / "prompt_chat_template.txt").write_text(prompt_text, encoding="utf-8")
    (chunk_dir / "raw_response.txt").write_text(raw_text, encoding="utf-8")
    _write_debug_json(
        chunk_dir / "meta.json",
        {
            "chunk_index": chunk_idx,
            "crop_xyxy": list(crop_xyxy) if crop_xyxy else None,
            "slot_count": len(infer_slots),
            "slot_ids": [s.get("slot_id") for s in infer_slots],
            "slots": infer_slots,
        },
    )
    _write_debug_json(chunk_dir / "parsed_chunk.json", chunk_data)
    _write_debug_json(chunk_dir / "normalized_pairs.json", {"pairs": normalized_pairs})


def run_inference(
    processor: Any,
    model: Any,
    model_device: Any,
    image_path: str,
    slots_json_path: str,
    output_json_path: str,
    max_new_tokens: int,
    temperature: float,
    slots_chunk_size: int,
    crop_by_slot_chunk: bool,
    crop_left_pad: int,
    crop_top_pad: int,
    crop_right_pad: int,
    crop_bottom_pad: int,
    max_image_size: int,
    resize_after_crop: bool,
    annotate_slot_id: bool,
    prompt_context_text: str = "",
    debug_dir: str = "",
) -> Dict[str, Any]:
    image = Image.open(image_path).convert("RGB")
    slots = extract_slots(slots_json_path)
    output_json = Path(output_json_path)
    dbg_dir = Path(debug_dir) if str(debug_dir or "").strip() else None
    if not slots:
        data = {"pairs": []}
        output_json.parent.mkdir(parents=True, exist_ok=True)
        output_json.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        if dbg_dir is not None:
            _write_debug_json(
                dbg_dir / "summary.json",
                {
                    "slot_count": 0,
                    "chunk_count": 0,
                    "pairs_count": 0,
                    "output_json": str(output_json),
                },
            )
        return data

    defer_resize_until_after_crop = bool(resize_after_crop)
    if (not defer_resize_until_after_crop) and max_image_size and max_image_size > 0:
        image, slots = _resize_image_and_slots_if_needed(image, slots, max_image_size)

    img_w, img_h = image.size
    slot_chunks = chunk_list(slots, int(slots_chunk_size))
    all_pairs: List[Dict[str, Any]] = []
    chunk_summaries: List[Dict[str, Any]] = []

    for chunk_idx, chunk in enumerate(slot_chunks, start=1):
        infer_image = image
        infer_slots = chunk
        crop_xyxy: Optional[Tuple[int, int, int, int]] = None
        if crop_by_slot_chunk:
            crop_xyxy = _build_crop_for_chunk(
                chunk,
                img_w,
                img_h,
                int(crop_left_pad),
                int(crop_top_pad),
                int(crop_right_pad),
                int(crop_bottom_pad),
            )
            cx1, cy1, cx2, cy2 = crop_xyxy
            infer_image = image.crop((cx1, cy1, cx2, cy2))
            infer_slots = _translate_chunk_slots_for_crop(chunk, crop_xyxy)

        if defer_resize_until_after_crop and max_image_size and max_image_size > 0:
            infer_image, infer_slots = _resize_image_and_slots_if_needed(infer_image, infer_slots, max_image_size)

        infer_image_for_model = _draw_slot_ids(infer_image, infer_slots) if annotate_slot_id else infer_image

        infer_w, infer_h = infer_image_for_model.size
        prompt = build_prompt(infer_slots, infer_w, infer_h, prompt_context_text=prompt_context_text)
        user_msg = [{"role": "user", "content": [{"type": "image"}, {"type": "text", "text": prompt}]}]
        prompt_text = processor.apply_chat_template(user_msg, tokenize=False, add_generation_prompt=True)
        inputs = processor(text=prompt_text, images=infer_image_for_model, return_tensors="pt").to(model_device)

        gen_kwargs = {"max_new_tokens": int(max_new_tokens), "do_sample": float(temperature) > 0.0}
        if float(temperature) > 0.0:
            gen_kwargs["temperature"] = float(temperature)

        with torch.no_grad():
            out = model.generate(**inputs, **gen_kwargs)

        generated = out[:, inputs["input_ids"].shape[1] :]
        text = processor.decode(generated[0], skip_special_tokens=True)
        try:
            chunk_data = extract_json_object(text)
        except Exception:
            chunk_data = json.loads(sanitize_json_text(text))
        normalized_pairs = _normalize_chunk_output(chunk_data, infer_slots)
        all_pairs.extend(normalized_pairs)
        _save_chunk_debug(
            debug_dir=dbg_dir,
            chunk_idx=chunk_idx,
            crop_xyxy=crop_xyxy,
            infer_slots=infer_slots,
            prompt=prompt,
            prompt_text=prompt_text,
            raw_text=text,
            chunk_data=chunk_data,
            normalized_pairs=normalized_pairs,
            infer_image=infer_image,
            infer_image_for_model=infer_image_for_model,
        )
        chunk_summaries.append(
            {
                "chunk_index": chunk_idx,
                "slot_count": len(infer_slots),
                "slot_ids": [s.get("slot_id") for s in infer_slots],
                "crop_xyxy": list(crop_xyxy) if crop_xyxy else None,
                "normalized_pair_count": len(normalized_pairs),
            }
        )

    by_sid: Dict[int, Dict[str, Any]] = {}
    grouped: Dict[int, List[Dict[str, Any]]] = defaultdict(list)
    for p in all_pairs:
        try:
            sid = int(p.get("slot_id"))
        except Exception:
            continue
        grouped[sid].append(p)

    for sid, plist in grouped.items():
        chosen = None
        for p in plist:
            key = (p.get("key") or {}).get("text")
            if p.get("is_valid_slot") and key:
                chosen = p
                break
        if chosen is None:
            chosen = plist[0]
        by_sid[sid] = chosen

    merged_pairs = [by_sid[k] for k in sorted(by_sid.keys())]
    data = {"pairs": merged_pairs}

    output_json.parent.mkdir(parents=True, exist_ok=True)
    output_json.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    if dbg_dir is not None:
        _write_debug_json(
            dbg_dir / "summary.json",
            {
                "image_path": image_path,
                "slots_json": slots_json_path,
                "output_json": str(output_json),
                "slot_count": len(slots),
                "chunk_count": len(slot_chunks),
                "pairs_count": len(merged_pairs),
                "slots_chunk_size": int(slots_chunk_size),
                "crop_by_slot_chunk": bool(crop_by_slot_chunk),
                "max_new_tokens": int(max_new_tokens),
                "temperature": float(temperature),
                "prompt_context_text": str(prompt_context_text or ""),
                "chunk_summaries": chunk_summaries,
            },
        )
        _write_debug_json(dbg_dir / "merged_pairs.json", data)
    return data


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--model_name", required=True, help="base model path or repo")
    ap.add_argument("--lora_path", required=True, help="LoRA adapter path")
    ap.add_argument("--image", required=True)
    ap.add_argument("--slots_json", required=True)
    ap.add_argument("--output_json", required=True)
    ap.add_argument("--max_new_tokens", type=int, default=20480)
    ap.add_argument("--temperature", type=float, default=0.0)
    ap.add_argument("--device_map", default="auto")
    ap.add_argument("--slots_chunk_size", type=int, default=8)
    ap.add_argument("--crop_by_slot_chunk", action="store_true")
    ap.add_argument("--crop_left_pad", type=int, default=220)
    ap.add_argument("--crop_top_pad", type=int, default=140)
    ap.add_argument("--crop_right_pad", type=int, default=80)
    ap.add_argument("--crop_bottom_pad", type=int, default=80)
    ap.add_argument("--max_image_size", type=int, default=1600)
    ap.add_argument("--resize_after_crop", action="store_true")
    ap.add_argument("--annotate_slot_id", action="store_true")
    ap.add_argument("--prompt_context_text", default="")
    ap.add_argument("--debug_dir", default="")
    args = ap.parse_args()

    processor = AutoProcessor.from_pretrained(args.model_name, trust_remote_code=True)
    dtype = torch.bfloat16 if torch.cuda.is_available() else torch.float32

    if Qwen3VLForConditionalGeneration is not None:
        base = Qwen3VLForConditionalGeneration.from_pretrained(
            args.model_name,
            dtype=dtype,
            trust_remote_code=True,
            device_map=args.device_map,
        )
    else:
        base = AutoModelForVision2Seq.from_pretrained(
            args.model_name,
            dtype=dtype,
            trust_remote_code=True,
            device_map=args.device_map,
        )

    model = PeftModel.from_pretrained(base, args.lora_path)
    model.eval()

    model_device = next(model.parameters()).device
    run_inference(
        processor=processor,
        model=model,
        model_device=model_device,
        image_path=args.image,
        slots_json_path=args.slots_json,
        output_json_path=args.output_json,
        max_new_tokens=int(args.max_new_tokens),
        temperature=float(args.temperature),
        slots_chunk_size=int(args.slots_chunk_size),
        crop_by_slot_chunk=bool(args.crop_by_slot_chunk),
        crop_left_pad=int(args.crop_left_pad),
        crop_top_pad=int(args.crop_top_pad),
        crop_right_pad=int(args.crop_right_pad),
        crop_bottom_pad=int(args.crop_bottom_pad),
        max_image_size=int(args.max_image_size),
        resize_after_crop=bool(args.resize_after_crop),
        annotate_slot_id=bool(args.annotate_slot_id),
        prompt_context_text=str(args.prompt_context_text or ""),
        debug_dir=str(args.debug_dir or ""),
    )


if __name__ == "__main__":
    main()
