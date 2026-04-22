import json
import os
import re
import sys
import subprocess
import threading
import base64
import mimetypes
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, Optional

import openpyxl
from openpyxl.utils import range_boundaries
try:
    from .lora_server import request_infer_lora_from_server
    from .slot_server import request_slot_detect_from_server
except ImportError:
    from lora_server import request_infer_lora_from_server  # type: ignore
    from slot_server import request_slot_detect_from_server  # type: ignore

ROOT = Path(__file__).resolve().parents[1]
VISION_DIR = ROOT / "vision"
SLOT_DETECT_PATH = VISION_DIR / "slot_detect.py"
LORA_INFER_PATH = VISION_DIR / "lora_infer.py"


def resolve_default_segformer_ckpt() -> Path:
    candidates = (
        VISION_DIR / "best_segformer.pt",
        Path.cwd() / "best.pt",
        Path.cwd() / "best_segformer_multitask.pt",
    )
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    return candidates[0]


DEFAULT_SEGFORMER_CKPT = resolve_default_segformer_ckpt()

if str(VISION_DIR) not in sys.path:
    sys.path.insert(0, str(VISION_DIR))

try:
    from excel_render import export_sheet_png_via_excel_com
except Exception as exc:
    export_sheet_png_via_excel_com = None
    _PIPELINE_IMPORT_ERROR = exc
else:
    _PIPELINE_IMPORT_ERROR = None


@dataclass
class PipelineParams:
    job_id: str
    job_dir: Path
    excel_path: Path
    sheet_name: Optional[str]
    dpi: int
    ckpt_path: Optional[str]
    qwen_base_url: Optional[str]
    qwen_model: Optional[str]
    qwen_api_key: Optional[str]
    qwen_mode: str
    qwen_lora_model: Optional[str]
    qwen_lora_path: Optional[str]
    qwen_lora_device_map: Optional[str]
    qwen_lora_max_new_tokens: int
    qwen_lora_temperature: float
    skip_qwen: bool
    skip_slots: bool
    use_llm: bool
    llm_base_url: Optional[str]
    llm_model: Optional[str]
    llm_api_key: Optional[str]
    input_text: str
    qwen_lora_slots_chunk_size: int = 8
    qwen_lora_crop_by_slot_chunk: bool = True
    qwen_lora_crop_left_pad: int = 220
    qwen_lora_crop_top_pad: int = 140
    qwen_lora_crop_right_pad: int = 80
    qwen_lora_crop_bottom_pad: int = 80
    qwen_lora_max_image_size: int = 1600
    qwen_lora_resize_after_crop: bool = True
    qwen_lora_annotate_slot_id: bool = True


class PipelineError(RuntimeError):
    pass


def _log(log_fn: Callable[[str], None], message: str) -> None:
    log_fn(message)


def _run_streaming(
    cmd: list[str],
    log_fn: Callable[[str], None],
    env: Optional[dict[str, str]] = None,
) -> int:
    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        bufsize=1,
        env=env,
    )

    def _pump(stream: Any) -> None:
        try:
            for chunk in iter(stream.readline, ""):
                line = chunk.rstrip("\r\n")
                if line:
                    _log(log_fn, line)
        finally:
            stream.close()

    threads = [
        threading.Thread(target=_pump, args=(proc.stdout,), daemon=True),
        threading.Thread(target=_pump, args=(proc.stderr,), daemon=True),
    ]
    for thread in threads:
        thread.start()
    rc = proc.wait()
    for thread in threads:
        thread.join()
    return int(rc)


def _encode_image_base64(image_path: Path) -> tuple[str, str]:
    mime, _ = mimetypes.guess_type(image_path.name)
    mime_token = (mime or "image/png").split("/")[-1]
    try:
        import io
        from PIL import Image
    except Exception:
        raw = image_path.read_bytes()
        return base64.b64encode(raw).decode("ascii"), mime_token

    max_dim = 7900
    with Image.open(image_path) as img:
        width, height = img.size
        if max(width, height) > max_dim:
            scale = min(max_dim / float(max(width, height)), 1.0)
            new_size = (
                max(1, int(width * scale)),
                max(1, int(height * scale)),
            )
            resample = getattr(Image, "Resampling", Image).LANCZOS
            img = img.resize(new_size, resample)
            if mime_token == "jpeg" and str(img.mode or "").upper() in {"RGBA", "LA", "P"}:
                img = img.convert("RGB")
            buf = io.BytesIO()
            save_format = "JPEG" if mime_token == "jpeg" else "PNG"
            save_mime = "jpeg" if mime_token == "jpeg" else "png"
            img.save(buf, format=save_format, optimize=True)
            return base64.b64encode(buf.getvalue()).decode("ascii"), save_mime

    raw = image_path.read_bytes()
    return base64.b64encode(raw).decode("ascii"), mime_token


def _extract_json_object(text: str) -> Dict[str, Any]:
    raw = str(text or "").strip()
    if not raw:
        return {}
    if "```" in raw:
        match = re.search(r"```(?:json)?\s*(.*?)```", raw, flags=re.IGNORECASE | re.DOTALL)
        if match:
            raw = match.group(1).strip()
    try:
        obj = json.loads(raw)
        return obj if isinstance(obj, dict) else {}
    except Exception:
        pass
    match = re.search(r"\{.*\}", raw, flags=re.DOTALL)
    if not match:
        return {}
    try:
        obj = json.loads(match.group(0))
        return obj if isinstance(obj, dict) else {}
    except Exception:
        return {}


def _build_slot_vlm_prompt(
    slots_batch: list[dict[str, Any]],
    *,
    img_w: int,
    img_h: int,
    prompt_context_text: str = "",
) -> str:
    lines: list[str] = []
    for slot in slots_batch:
        slot_id = int(slot.get("slot_id", 0) or 0)
        bbox = slot.get("pred_bbox_px") or [-1, -1, -1, -1]
        value_cell = str(slot.get("value_range_a1") or "").strip()
        lines.append(f'{slot_id}: bbox={bbox} value_cell="{value_cell}"')
    slot_block = "\n".join(lines)
    context_block = ""
    if str(prompt_context_text or "").strip():
        context_block = (
            "Additional local matching constraints:\n"
            f"{str(prompt_context_text).strip()}\n\n"
        )
    return (
        "You are given a spreadsheet or form screenshot.\n"
        "The red translucent boxes, yellow slot ids, highlights, and outlines were added by the system as overlay markers. "
        "Ignore them as worksheet content. They are not keys, labels, or values.\n"
        "For each slot, decide whether it is a real writable field (`is_valid_slot`) and identify the most likely key/label text for that slot.\n"
        "Prefer the nearest structural label in this order: same row on the left, header above, row header on the left, then other clearly associated nearby label text.\n"
        "Do not treat long notes or paragraphs as a key unless they directly label the slot.\n"
        "Copy `value_cell` exactly from the provided slot list.\n"
        "If no reliable key is visible, return `is_valid_slot=false`, an empty key text, bbox `[-1,-1,-1,-1]`, and a short reason.\n\n"
        f"{context_block}"
        f"Image size: {img_w}x{img_h}\n\n"
        "Slots to process:\n"
        f"{slot_block}\n\n"
        "Return JSON only with schema:\n"
        "{"
        "\"pairs\":["
        "{"
        "\"slot_id\":1,"
        "\"value_cell\":\"B2\","
        "\"is_valid_slot\":true,"
        "\"key\":{\"text\":\"Name\",\"bbox\":[0,0,10,10]},"
        "\"confidence\":0.9,"
        "\"reason\":\"same-row left label\""
        "}"
        "]"
        "}\n"
        "Constraints:\n"
        "- Every slot_id must appear exactly once.\n"
        "- confidence must be a number between 0 and 1.\n"
        "- Output JSON only, no markdown."
    )


def run_slot_vlm(
    image_path: Path,
    slots_json: Path,
    out_dir: Path,
    model_name: str,
    base_url: Optional[str],
    api_key: Optional[str],
    temperature: float,
    log_fn: Callable[[str], None],
    prompt_context_text: str = "",
    batch_size: int = 8,
    debug_dir: Path | None = None,
) -> Path:
    from openai import OpenAI

    if not image_path.exists():
        raise PipelineError(f"Sheet image not found: {image_path}")
    if not slots_json.exists():
        raise PipelineError(f"slots_json not found: {slots_json}")
    model_name = str(model_name or "").strip()
    if not model_name:
        raise PipelineError("agent_model is required for slot VLM inference")

    payload = json.loads(slots_json.read_text(encoding="utf-8"))
    slots = payload.get("slots") if isinstance(payload, dict) else None
    if not isinstance(slots, list) or not slots:
        raise PipelineError(f"No slots found in {slots_json}")
    image_size = payload.get("image_size") if isinstance(payload, dict) else {}
    img_w = int((image_size or {}).get("w", 0) or 0)
    img_h = int((image_size or {}).get("h", 0) or 0)
    if img_w <= 0 or img_h <= 0:
        try:
            from PIL import Image

            with Image.open(image_path) as img:
                img_w, img_h = img.size
        except Exception:
            img_w = 0
            img_h = 0

    out_dir.mkdir(parents=True, exist_ok=True)
    if debug_dir is not None:
        debug_dir.mkdir(parents=True, exist_ok=True)
    output_json = out_dir / "qwen_pairs.json"
    client = OpenAI(
        api_key=api_key or os.environ.get("OPENAI_API_KEY") or os.environ.get("PPCHAT_API_KEY"),
        base_url=base_url or os.environ.get("OPENAI_BASE_URL"),
        timeout=3600,
    )
    image_b64, image_mime = _encode_image_base64(image_path)
    _log(log_fn, f"Running slot VLM inference via agent model {model_name} ...")

    all_pairs: list[dict[str, Any]] = []
    batch_size = max(1, int(batch_size or 1))
    for batch_index, start in enumerate(range(0, len(slots), batch_size)):
        batch = slots[start : start + batch_size]
        prompt = _build_slot_vlm_prompt(
            batch,
            img_w=img_w,
            img_h=img_h,
            prompt_context_text=prompt_context_text,
        )
        messages = [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url", "image_url": {"url": f"data:image/{image_mime};base64,{image_b64}"}},
                ],
            }
        ]
        try:
            resp = client.chat.completions.create(
                model=model_name,
                messages=messages,
                temperature=float(temperature),
                response_format={"type": "json_object"},
            )
        except Exception:
            resp = client.chat.completions.create(
                model=model_name,
                messages=messages,
                temperature=float(temperature),
            )
        raw = str(resp.choices[0].message.content or "").strip()
        if debug_dir is not None:
            (debug_dir / f"slot_vlm_raw_{batch_index:03d}.txt").write_text(raw, encoding="utf-8")
        data = _extract_json_object(raw)
        pairs = data.get("pairs") if isinstance(data, dict) else None
        if not isinstance(pairs, list):
            pairs = []
        existing: dict[int, dict[str, Any]] = {}
        for item in pairs:
            if not isinstance(item, dict) or not str(item.get("slot_id", "")).strip():
                continue
            try:
                existing[int(item.get("slot_id", -1))] = item
            except Exception:
                continue
        for slot in batch:
            slot_id = int(slot.get("slot_id", 0) or 0)
            value_cell = str(slot.get("value_range_a1") or "").strip()
            item = existing.get(slot_id)
            if item is None:
                all_pairs.append(
                    {
                        "slot_id": slot_id,
                        "value_cell": value_cell,
                        "is_valid_slot": False,
                        "key": {"text": "", "bbox": [-1, -1, -1, -1]},
                        "confidence": 0.0,
                        "reason": "missing slot_id in model output",
                    }
                )
                continue
            normalized = dict(item)
            normalized["slot_id"] = slot_id
            normalized["value_cell"] = value_cell
            key_obj = normalized.get("key")
            if not isinstance(key_obj, dict):
                key_obj = {}
            bbox = key_obj.get("bbox")
            if not (isinstance(bbox, list) and len(bbox) == 4):
                bbox = [-1, -1, -1, -1]
            key_obj["bbox"] = bbox
            key_obj["text"] = str(key_obj.get("text") or "").strip()
            normalized["key"] = key_obj
            normalized["is_valid_slot"] = bool(normalized.get("is_valid_slot"))
            try:
                normalized["confidence"] = float(normalized.get("confidence", 0.0) or 0.0)
            except Exception:
                normalized["confidence"] = 0.0
            normalized["reason"] = str(normalized.get("reason") or "").strip()
            all_pairs.append(normalized)

    all_pairs.sort(key=lambda item: int(item.get("slot_id", 10**9) or 10**9))
    output_json.write_text(json.dumps({"pairs": all_pairs}, ensure_ascii=False, indent=2), encoding="utf-8")
    return output_json


def ensure_xlsx(src: Path, dst: Path, log_fn: Callable[[str], None]) -> Path:
    if src.suffix.lower() == ".xlsx":
        if src.resolve() != dst.resolve():
            dst.write_bytes(src.read_bytes())
        return dst
    if src.suffix.lower() != ".xls":
        raise PipelineError(f"Unsupported file type: {src.suffix}")

    _log(log_fn, "Converting .xls to .xlsx via Excel COM ...")
    try:
        import pythoncom
        pythoncom.CoInitialize()
    except Exception:
        pythoncom = None

    try:
        import win32com.client as win32
    except Exception as exc:
        raise PipelineError("pywin32 is required for Excel COM conversion") from exc

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
            str(src.resolve()),
            UpdateLinks=0,
            ReadOnly=True,
            AddToMru=False,
        )
        # 51 = xlOpenXMLWorkbook (.xlsx)
        wb.SaveAs(str(dst.resolve()), FileFormat=51)
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
        if pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    if not dst.exists():
        raise PipelineError("Excel COM conversion failed: output file not created")

    return dst


def preprocess_excel(
    xlsx_path: Path,
    out_dir: Path,
    sheet_name: Optional[str],
    dpi: int,
    log_fn: Callable[[str], None],
) -> Dict[str, Path]:
    if export_sheet_png_via_excel_com is None:
        raise PipelineError(f"excel_render import failed: {_PIPELINE_IMPORT_ERROR}")

    out_dir.mkdir(parents=True, exist_ok=True)
    sheet_png = out_dir / "sheet.png"
    bounds_json = out_dir / "sheet_bounds.json"
    edges_json = out_dir / "sheet_edges.json"

    _log(log_fn, "Exporting sheet image via Excel COM ...")
    try:
        import pythoncom
        pythoncom.CoInitialize()
    except Exception:
        pythoncom = None

    try:
        min_row, min_col, max_row, max_col, x_edges, y_edges = export_sheet_png_via_excel_com(
            str(xlsx_path),
            str(sheet_png),
            sheet_name=sheet_name,
            dpi=dpi,
        )
    finally:
        if pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    with open(bounds_json, "w", encoding="utf-8") as f:
        json.dump(
            {
                "min_row": int(min_row),
                "min_col": int(min_col),
                "max_row": int(max_row),
                "max_col": int(max_col),
            },
            f,
            ensure_ascii=False,
            indent=2,
        )

    with open(edges_json, "w", encoding="utf-8") as f:
        json.dump(
            {"dpi": int(dpi), "x_edges": x_edges, "y_edges": y_edges},
            f,
            ensure_ascii=False,
        )

    return {
        "sheet_png": sheet_png,
        "bounds_json": bounds_json,
        "edges_json": edges_json,
    }


def run_qwen_small(
    image_path: Path,
    xlsx_path: Path,
    out_dir: Path,
    sheet_name: Optional[str],
    ckpt_path: Optional[str],
    edges_json: Optional[Path],
    bounds_json: Optional[Path],
    qwen_base_url: Optional[str],
    qwen_model: Optional[str],
    qwen_api_key: Optional[str],
    skip_qwen: bool,
    log_fn: Callable[[str], None],
) -> Dict[str, Path]:
    if not SLOT_DETECT_PATH.exists():
        raise PipelineError(f"Missing slot_detect.py at {SLOT_DETECT_PATH}")

    ckpt_to_use: Optional[Path] = None
    if ckpt_path:
        ckpt_to_use = Path(ckpt_path)
        if not ckpt_to_use.exists():
            raise PipelineError(f"Checkpoint not found: {ckpt_to_use}")
    else:
        default_ckpt = resolve_default_segformer_ckpt()
        if not default_ckpt.exists():
            raise PipelineError(
                "Segformer checkpoint not found. Provide --ckpt_path in the UI or place a compatible best.pt/"
                f"best_segformer.pt under one of the expected locations, e.g. {default_ckpt}."
            )
        ckpt_to_use = default_ckpt

    out_dir.mkdir(parents=True, exist_ok=True)

    reuse_addr = str(os.getenv("SLOT_DETECT_REUSE_ADDR") or "").strip()
    if reuse_addr:
        _log(log_fn, f"Running slot_detect.py via preloaded slot server {reuse_addr} ...")
        try:
            return request_slot_detect_from_server(
                address=reuse_addr,
                image_path=image_path,
                xlsx_path=xlsx_path,
                out_dir=out_dir,
                sheet_name=sheet_name,
                edges_json=edges_json,
                bounds_json=bounds_json,
                qwen_base_url=qwen_base_url or "",
                qwen_model=qwen_model or "",
                qwen_api_key=qwen_api_key or "",
                skip_qwen=skip_qwen,
            )
        except Exception as exc:
            _log(log_fn, f"[WARN] slot-detect reuse server failed, falling back to subprocess: {exc}")

    cmd = [
        sys.executable,
        str(SLOT_DETECT_PATH),
        "--image",
        str(image_path),
        "--xlsx",
        str(xlsx_path),
        "--out_dir",
        str(out_dir),
    ]

    if sheet_name:
        cmd += ["--sheet", sheet_name]
    if ckpt_to_use:
        cmd += ["--ckpt", str(ckpt_to_use)]
    if edges_json:
        cmd += ["--edges_json", str(edges_json)]
    if bounds_json:
        cmd += ["--bounds_json", str(bounds_json)]
    if skip_qwen:
        cmd.append("--skip_qwen")

    env = os.environ.copy()
    if qwen_base_url:
        env["QWEN_BASE_URL"] = qwen_base_url
    if qwen_model:
        env["QWEN_MODEL"] = qwen_model
    if qwen_api_key:
        env["QWEN_API_KEY"] = qwen_api_key

    _log(log_fn, "Running slot_detect.py (slot detection + alignment) ...")
    rc = _run_streaming(
        cmd,
        log_fn=log_fn,
        env=env,
    )
    if rc != 0:
        raise PipelineError("slot_detect.py failed, check logs for details")

    outputs = {
        "mask_png": out_dir / "mask.png",
        "overlay_png": out_dir / "overlay.png",
        "slots_json": out_dir / "slots_aligned.json",
        "qwen_pairs": out_dir / "qwen_pairs.json",
    }

    return outputs


def run_qwen_lora(
    image_path: Path,
    slots_json: Path,
    out_dir: Path,
    model_name: str,
    lora_path: str,
    device_map: str,
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
    log_fn: Callable[[str], None],
    prompt_context_text: str = "",
    debug_dir: Path | None = None,
) -> Path:
    if not LORA_INFER_PATH.exists():
        raise PipelineError(f"Missing lora_infer.py at {LORA_INFER_PATH}")

    out_dir.mkdir(parents=True, exist_ok=True)
    output_json = out_dir / "qwen_pairs.json"
    reuse_addr = str(os.getenv("QWEN_LORA_REUSE_ADDR") or "").strip()
    if reuse_addr:
        _log(log_fn, f"Running lora_infer.py via preloaded LoRA server {reuse_addr} ...")
        try:
            request_infer_lora_from_server(
                address=reuse_addr,
                image_path=image_path,
                slots_json=slots_json,
                output_json=output_json,
                max_new_tokens=max_new_tokens,
                temperature=temperature,
                slots_chunk_size=slots_chunk_size,
                crop_by_slot_chunk=crop_by_slot_chunk,
                crop_left_pad=crop_left_pad,
                crop_top_pad=crop_top_pad,
                crop_right_pad=crop_right_pad,
                crop_bottom_pad=crop_bottom_pad,
                max_image_size=max_image_size,
                resize_after_crop=resize_after_crop,
                annotate_slot_id=annotate_slot_id,
                prompt_context_text=prompt_context_text,
                debug_dir=debug_dir,
            )
        except Exception as exc:
            _log(log_fn, f"[WARN] Qwen LoRA reuse server failed, falling back to subprocess: {exc}")
        else:
            if not output_json.exists():
                raise PipelineError("Qwen LoRA reuse server did not produce qwen_pairs.json")
            return output_json

    cmd = [
        sys.executable,
        str(LORA_INFER_PATH),
        "--model_name",
        model_name,
        "--lora_path",
        lora_path,
        "--image",
        str(image_path),
        "--slots_json",
        str(slots_json),
        "--output_json",
        str(output_json),
        "--max_new_tokens",
        str(max_new_tokens),
        "--temperature",
        str(temperature),
        "--device_map",
        device_map,
        "--slots_chunk_size",
        str(slots_chunk_size),
        "--crop_left_pad",
        str(crop_left_pad),
        "--crop_top_pad",
        str(crop_top_pad),
        "--crop_right_pad",
        str(crop_right_pad),
        "--crop_bottom_pad",
        str(crop_bottom_pad),
        "--max_image_size",
        str(max_image_size),
    ]
    if str(prompt_context_text or "").strip():
        cmd.extend(["--prompt_context_text", str(prompt_context_text)])
    if debug_dir is not None:
        cmd.extend(["--debug_dir", str(debug_dir)])
    if crop_by_slot_chunk:
        cmd.append("--crop_by_slot_chunk")
    if resize_after_crop:
        cmd.append("--resize_after_crop")
    if annotate_slot_id:
        cmd.append("--annotate_slot_id")

    _log(log_fn, "Running lora_infer.py (local Qwen LoRA) ...")
    rc = _run_streaming(
        cmd,
        log_fn=log_fn,
    )
    if rc != 0:
        raise PipelineError("lora_infer.py failed, check logs for details")

    if not output_json.exists():
        raise PipelineError("lora_infer.py did not produce qwen_pairs.json")

    return output_json


def build_form_from_qwen(
    qwen_pairs_path: Path,
    xlsx_path: Optional[Path] = None,
    sheet_name: Optional[str] = None,
) -> Dict[str, Any]:
    return _build_form_from_qwen_impl(qwen_pairs_path, xlsx_path, sheet_name)


def _normalize_a1_range_local(a1: str) -> str:
    s = str(a1 or "").strip().upper().replace(" ", "")
    if not s:
        raise ValueError("empty range")
    if ":" not in s:
        _ = range_boundaries(s)
        return f"{s}:{s}"
    c1, r1, c2, r2 = range_boundaries(s)
    min_c, max_c = min(c1, c2), max(c1, c2)
    min_r, max_r = min(r1, r2), max(r1, r2)
    left = f"{openpyxl.utils.get_column_letter(min_c)}{min_r}"
    right = f"{openpyxl.utils.get_column_letter(max_c)}{max_r}"
    return f"{left}:{right}"


def _iter_range_cells_local(range_str: str) -> list[str]:
    out: list[str] = []
    for token in [x.strip() for x in str(range_str or "").split(",") if x.strip()]:
        c1, r1, c2, r2 = range_boundaries(_normalize_a1_range_local(token))
        for rr in range(r1, r2 + 1):
            for cc in range(c1, c2 + 1):
                out.append(f"{openpyxl.utils.get_column_letter(cc)}{rr}")
    return out


def _merged_anchor_map_local(ws) -> dict[tuple[int, int], tuple[int, int]]:
    out: dict[tuple[int, int], tuple[int, int]] = {}
    for mr in ws.merged_cells.ranges:
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                out[(r, c)] = (mr.min_row, mr.min_col)
    return out


def _pair_points_to_own_label(
    ws,
    merge_anchor: dict[tuple[int, int], tuple[int, int]],
    key_text: str,
    value_cell: str,
) -> bool:
    nk = normalize_key(key_text)
    if not nk:
        return False
    for cell_ref in _iter_range_cells_local(value_cell):
        try:
            min_col, min_row, _, _ = range_boundaries(_normalize_a1_range_local(cell_ref))
        except Exception:
            continue
        anchor = merge_anchor.get((min_row, min_col), (min_row, min_col))
        value = ws.cell(row=anchor[0], column=anchor[1]).value
        if normalize_key(value) == nk:
            return True
    return False


def _build_form_from_qwen_impl(
    qwen_pairs_path: Path,
    xlsx_path: Optional[Path],
    sheet_name: Optional[str],
) -> Dict[str, Any]:
    with open(qwen_pairs_path, "r", encoding="utf-8") as f:
        payload = json.load(f)

    wb = None
    ws = None
    merge_anchor: dict[tuple[int, int], tuple[int, int]] = {}
    if xlsx_path is not None:
        wb = openpyxl.load_workbook(xlsx_path, data_only=True)
        ws = wb[sheet_name] if (sheet_name and sheet_name in wb.sheetnames) else wb.active
        merge_anchor = _merged_anchor_map_local(ws)

    try:
        pairs = []
        invalid_slots = []
        dropped_suspicious_pairs = []
        payload_modified = False
        for item in payload.get("pairs", []) or []:
            if not isinstance(item, dict):
                continue
            key = (item.get("key") or {}).get("text")
            key_text = str(key).strip() if key is not None else ""
            if not key:
                key_text = ""
            value_cell = str(item.get("value_cell") or "").strip()
            raw_reason = str(
                item.get("pipeline_invalidated_reason")
                or item.get("reason")
                or ""
            ).strip()

            def _append_invalid_slot(reason: str) -> None:
                if not value_cell:
                    return
                invalid_slots.append(
                    {
                        "slot_id": item.get("slot_id"),
                        "key": key_text,
                        "value_cell": value_cell,
                        "reason": reason,
                        "pipeline_invalidated": bool(item.get("pipeline_invalidated")),
                    }
                )

            if not item.get("is_valid_slot"):
                _append_invalid_slot(raw_reason or "qwen_marked_invalid_slot")
                continue
            if not key_text:
                _append_invalid_slot(raw_reason or "missing_key_text")
                continue
            if not value_cell:
                continue
            if ws is not None and _pair_points_to_own_label(ws, merge_anchor, str(key).strip(), value_cell):
                item["is_valid_slot"] = False
                invalid_reason = "pipeline_invalid_self_label_value_cell"
                existing_reason = str(item.get("reason") or "").strip()
                if existing_reason:
                    reason_tokens = {tok.strip() for tok in existing_reason.split(";") if tok.strip()}
                    if invalid_reason not in reason_tokens:
                        item["reason"] = f"{existing_reason}; {invalid_reason}"
                else:
                    item["reason"] = invalid_reason
                item["pipeline_invalidated"] = True
                item["pipeline_invalidated_reason"] = invalid_reason
                payload_modified = True
                _append_invalid_slot(invalid_reason)
                dropped_suspicious_pairs.append(
                    {
                        "slot_id": item.get("slot_id"),
                        "key": key_text,
                        "value_cell": value_cell,
                        "reason": "value_cell currently contains the same text as the key/label",
                    }
                )
                continue
            pairs.append(
                {
                    "slot_id": item.get("slot_id"),
                    "key": key_text,
                    "value_cell": value_cell,
                }
            )
        if payload_modified:
            with open(qwen_pairs_path, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
        dedup_invalid_slots = []
        seen_invalid = set()
        for item in invalid_slots:
            sig = (
                str(item.get("slot_id") or ""),
                str(item.get("value_cell") or "").upper(),
                normalize_key(str(item.get("key") or "")),
                str(item.get("reason") or ""),
            )
            if sig in seen_invalid:
                continue
            seen_invalid.add(sig)
            dedup_invalid_slots.append(item)
        return {
            "pairs": pairs,
            "invalid_slots": dedup_invalid_slots,
            "dropped_suspicious_pairs": dropped_suspicious_pairs,
        }
    finally:
        if wb is not None:
            wb.close()


def normalize_key(text: str) -> str:
    if text is None:
        return ""
    s = str(text).strip().lower()
    if s.endswith(":"):
        s = s[:-1].strip()
    # convert underscores and punctuation to spaces
    s = re.sub(r"[^a-z0-9]+", " ", s)
    s = " ".join(s.split())
    return s


def parse_kv_text(raw: str) -> Dict[str, Any]:
    text = raw.strip()
    if not text:
        return {}

    if text.startswith("{") and text.endswith("}"):
        try:
            data = json.loads(text)
            if isinstance(data, dict):
                return data
        except json.JSONDecodeError:
            pass

    if text.startswith("[") and text.endswith("]"):
        try:
            data = json.loads(text)
            if isinstance(data, list):
                out = {}
                for item in data:
                    if isinstance(item, dict):
                        key = item.get("key") or item.get("name")
                        value = item.get("value")
                        if key is not None:
                            out[str(key).strip()] = value
                return out
        except json.JSONDecodeError:
            pass

    out: Dict[str, Any] = {}
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        if line.startswith("#"):
            continue
        if ":" in line:
            key, value = line.split(":", 1)
        elif "=" in line:
            key, value = line.split("=", 1)
        else:
            continue
        key = key.strip()
        value = value.strip()
        if key:
            out[key] = value
    return out


def parse_kv_llm(
    text: str,
    model: str,
    base_url: Optional[str],
    api_key: Optional[str],
) -> Dict[str, Any]:
    from openai import OpenAI

    client = OpenAI(api_key=api_key or os.environ.get("OPENAI_API_KEY"), base_url=base_url)
    schema_keys = [
        "hotel_name",
        "check_in_date",
        "check_out_date",
        "guest_1_name",
        "guest_1_name_spelling",
        "guest_1_bed_type",
        "guest_1_smoking_preference",
        "guest_1_confirmation_number",
        "guest_2_name",
        "guest_2_name_spelling",
        "guest_2_bed_type",
        "guest_2_smoking_preference",
        "guest_2_confirmation_number",
        "guest_3_name",
        "guest_3_name_spelling",
        "guest_3_bed_type",
        "guest_3_smoking_preference",
        "guest_3_confirmation_number",
        "fhsu_contact_name",
        "fhsu_department",
        "fhsu_college",
        "fhsu_fax_number",
        "fhsu_phone_number",
        "bpc_last_5_digits",
        "will_pay_personally",
        "booking_type",
        "fax_to_number",
        "note",
    ]
    schema_block = "\n".join(schema_keys)
    prompt = (
        "Extract fields and return ONLY a valid JSON object with EXACTLY the keys below.\n"
        "Rules:\n"
        "- Do not add or remove keys.\n"
        "- If a value is missing, use \"\".\n"
        "- Preserve the original casing and wording from the input text for values.\n"
        "- Output must be a single JSON object, no extra text.\n\n"
        "Keys:\n"
        f"{schema_block}\n"
    )

    resp = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": "You are a strict JSON generator."},
            {"role": "user", "content": prompt + "\n\n" + text},
        ],
        temperature=0.0,
        response_format={"type": "json_object"},
    )
    content = resp.choices[0].message.content
    return json.loads(content)


def fill_excel(
    template_xlsx: Path,
    form: Dict[str, Any],
    kv: Dict[str, Any],
    out_path: Path,
    sheet_name: Optional[str],
) -> Dict[str, Any]:
    wb = openpyxl.load_workbook(template_xlsx)
    ws = wb[sheet_name] if sheet_name else wb.active

    # Build merged-cell anchor map: any cell in a merged range -> top-left cell
    merge_anchor = {}
    for mr in ws.merged_cells.ranges:
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                merge_anchor[(r, c)] = (mr.min_row, mr.min_col)

    # normalize kv keys for robust matching
    kv_norm: Dict[str, Any] = {}
    for k, v in (kv or {}).items():
        nk = normalize_key(k)
        if nk and nk not in kv_norm:
            kv_norm[nk] = v

    filled_pairs = []
    for pair in form.get("pairs", []) or []:
        key = pair.get("key")
        cell_range = pair.get("value_cell")
        if not key or not cell_range:
            continue
        nk = normalize_key(key)
        if not nk or nk not in kv_norm:
            continue
        value = kv_norm[nk]
        try:
            min_col, min_row, _, _ = range_boundaries(cell_range)
        except Exception:
            continue
        anchor = merge_anchor.get((min_row, min_col), (min_row, min_col))
        ws.cell(row=anchor[0], column=anchor[1]).value = value
        filled_pairs.append({"key": key, "value": value, "value_cell": cell_range})

    wb.save(out_path)
    return {"pairs": filled_pairs}


def run_pipeline(params: PipelineParams, log_fn: Callable[[str], None]) -> Dict[str, Path]:
    outputs: Dict[str, Path] = {}

    if not params.ckpt_path and DEFAULT_SEGFORMER_CKPT.exists():
        params.ckpt_path = str(DEFAULT_SEGFORMER_CKPT)
        _log(log_fn, f"Using default SegFormer ckpt: {params.ckpt_path}")

    params.job_dir.mkdir(parents=True, exist_ok=True)
    xlsx_path = params.job_dir / "input.xlsx"
    xlsx_path = ensure_xlsx(params.excel_path, xlsx_path, log_fn)
    outputs["input_xlsx"] = xlsx_path

    # Resolve a safe sheet name if not provided or invalid.
    sheet_name = params.sheet_name
    try:
        wb = openpyxl.load_workbook(xlsx_path, read_only=True)
        if not wb.sheetnames:
            raise PipelineError("Workbook contains no worksheets.")
        if sheet_name and sheet_name not in wb.sheetnames:
            raise PipelineError(f"Sheet not found: {sheet_name}. Available: {', '.join(wb.sheetnames)}")
        if not sheet_name:
            sheet_name = wb.sheetnames[0]
    finally:
        try:
            wb.close()
        except Exception:
            pass

    preprocess_dir = params.job_dir / "preprocess"
    preprocess_outputs = preprocess_excel(
        xlsx_path=xlsx_path,
        out_dir=preprocess_dir,
        sheet_name=sheet_name,
        dpi=params.dpi,
        log_fn=log_fn,
    )
    outputs.update(preprocess_outputs)

    if params.skip_slots:
        return outputs

    qwen_out_dir = params.job_dir / "qwen_out"
    qwen_outputs = run_qwen_small(
        image_path=preprocess_outputs["sheet_png"],
        xlsx_path=xlsx_path,
        out_dir=qwen_out_dir,
        sheet_name=sheet_name,
        ckpt_path=params.ckpt_path,
        edges_json=preprocess_outputs["edges_json"],
        bounds_json=preprocess_outputs["bounds_json"],
        qwen_base_url=params.qwen_base_url,
        qwen_model=params.qwen_model,
        qwen_api_key=params.qwen_api_key,
        skip_qwen=(params.skip_qwen or params.qwen_mode == "lora"),
        log_fn=log_fn,
    )
    outputs.update(qwen_outputs)

    if params.skip_qwen:
        return outputs

    qwen_pairs_path: Optional[Path] = None
    if params.qwen_mode == "lora":
        if not params.qwen_lora_model or not params.qwen_lora_path:
            raise PipelineError("qwen_lora_model and qwen_lora_path are required for qwen_mode=lora")
        qwen_pairs_path = run_qwen_lora(
            image_path=preprocess_outputs["sheet_png"],
            slots_json=qwen_outputs["slots_json"],
            out_dir=qwen_out_dir,
            model_name=params.qwen_lora_model,
            lora_path=params.qwen_lora_path,
            device_map=params.qwen_lora_device_map or "auto",
            max_new_tokens=params.qwen_lora_max_new_tokens,
            temperature=params.qwen_lora_temperature,
            slots_chunk_size=params.qwen_lora_slots_chunk_size,
            crop_by_slot_chunk=params.qwen_lora_crop_by_slot_chunk,
            crop_left_pad=params.qwen_lora_crop_left_pad,
            crop_top_pad=params.qwen_lora_crop_top_pad,
            crop_right_pad=params.qwen_lora_crop_right_pad,
            crop_bottom_pad=params.qwen_lora_crop_bottom_pad,
            max_image_size=params.qwen_lora_max_image_size,
            resize_after_crop=params.qwen_lora_resize_after_crop,
            annotate_slot_id=params.qwen_lora_annotate_slot_id,
            log_fn=log_fn,
            debug_dir=qwen_out_dir / "lora_debug",
        )
        outputs["qwen_pairs"] = qwen_pairs_path
    else:
        if qwen_outputs["qwen_pairs"].exists():
            qwen_pairs_path = qwen_outputs["qwen_pairs"]
        else:
            raise PipelineError("qwen_pairs.json not found; cannot build form.json")

    if not qwen_pairs_path:
        raise PipelineError("qwen_pairs.json not found; cannot build form.json")

    form = build_form_from_qwen(qwen_pairs_path, xlsx_path=xlsx_path, sheet_name=sheet_name)
    dropped_suspicious = form.get("dropped_suspicious_pairs", []) if isinstance(form, dict) else []
    if dropped_suspicious:
        _log(log_fn, f"[INFO] dropped suspicious qwen pairs: {len(dropped_suspicious)}")
    form_path = params.job_dir / "form.json"
    with open(form_path, "w", encoding="utf-8") as f:
        json.dump(form, f, ensure_ascii=False, indent=2)
    outputs["form_json"] = form_path

    if params.use_llm:
        if not params.llm_model:
            raise PipelineError("LLM model name is required when use_llm is enabled")
        kv = parse_kv_llm(params.input_text, params.llm_model, params.llm_base_url, params.llm_api_key)
    else:
        kv = parse_kv_text(params.input_text)

    filled_json_path = params.job_dir / "filled.json"
    with open(filled_json_path, "w", encoding="utf-8") as f:
        json.dump(kv, f, ensure_ascii=False, indent=2)
    outputs["filled_json"] = filled_json_path

    filled_xlsx_path = params.job_dir / "filled.xlsx"
    filled_pairs = fill_excel(
        template_xlsx=xlsx_path,
        form=form,
        kv=kv,
        out_path=filled_xlsx_path,
        sheet_name=params.sheet_name,
    )
    outputs["filled_xlsx"] = filled_xlsx_path

    filled_pairs_path = params.job_dir / "filled_pairs.json"
    with open(filled_pairs_path, "w", encoding="utf-8") as f:
        json.dump(filled_pairs, f, ensure_ascii=False, indent=2)
    outputs["filled_pairs"] = filled_pairs_path

    return outputs
