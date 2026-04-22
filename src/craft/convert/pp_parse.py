#!/usr/bin/env python3
from __future__ import annotations

"""Parse form-like PDFs/images with PP-StructureV3 and export structured outputs.

Example:
    python -m craft.convert.pp_parse ^
        --input_path ./samples ^
        --recursive ^
        --save_debug_images
"""

import argparse
import hashlib
import inspect
import json
import sys
from collections.abc import Mapping
from pathlib import Path
from typing import Any


IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".bmp", ".webp", ".tif", ".tiff"}
PDF_SUFFIXES = {".pdf"}


def _ensure_dir(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path


def _sanitize_name(name: str) -> str:
    import re

    safe = re.sub(r"[^\w.-]+", "_", str(name or "").strip(), flags=re.UNICODE)
    safe = safe.strip("._")
    return safe or "document"


def _doc_id(path: Path) -> str:
    digest = hashlib.md5(str(path.resolve()).encode("utf-8"), usedforsecurity=False).hexdigest()[:8]
    return f"{_sanitize_name(path.stem)}_{digest}"


def discover_inputs(input_path: Path, recursive: bool) -> list[Path]:
    if input_path.is_file():
        if input_path.suffix.lower() not in (IMAGE_SUFFIXES | PDF_SUFFIXES):
            raise ValueError(f"Unsupported input file type: {input_path}")
        return [input_path]
    if not input_path.is_dir():
        raise FileNotFoundError(f"Input path not found: {input_path}")
    iterator = input_path.rglob("*") if recursive else input_path.iterdir()
    docs = [
        p
        for p in iterator
        if p.is_file() and p.suffix.lower() in (IMAGE_SUFFIXES | PDF_SUFFIXES)
    ]
    return sorted(docs)

def _json_ready(value: Any) -> Any:
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, Path):
        return str(value)
    if hasattr(value, "tolist"):
        return _json_ready(value.tolist())
    if isinstance(value, bytes):
        return {
            "__type__": "bytes",
            "length": len(value),
        }
    if isinstance(value, Mapping):
        return {str(k): _json_ready(v) for k, v in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [_json_ready(v) for v in value]
    if hasattr(value, "mode") and hasattr(value, "size") and value.__class__.__name__.lower().endswith("image"):
        return {
            "__type__": value.__class__.__name__,
            "mode": getattr(value, "mode", None),
            "size": list(getattr(value, "size", []) or []),
        }
    if hasattr(value, "__dict__"):
        return {
            "__type__": value.__class__.__name__,
            "repr": repr(value),
        }
    return str(value)



def init_pipeline(args: argparse.Namespace) -> Any:
    try:
        from paddleocr import PPStructureV3  # type: ignore
    except Exception as exc:  # noqa: BLE001
        raise RuntimeError(
            "Missing PP-StructureV3 dependencies. Install PaddlePaddle first, then PaddleOCR.\n"
            "Example:\n"
            "  python -m pip install paddleocr pymupdf\n"
            "PP-StructureV3 is part of PaddleOCR 3.x."
        ) from exc

    sig = inspect.signature(PPStructureV3.__init__)
    params = sig.parameters
    kwargs: dict[str, Any] = {}

    candidate_kwargs = {
        "device": args.device or None,
        "use_doc_orientation_classify": args.use_doc_orientation_classify,
        "use_doc_unwarping": args.use_doc_unwarping,
        "use_table_recognition": args.use_table_recognition,
        "use_seal_recognition": args.use_seal_recognition,
        "use_formula_recognition": args.use_formula_recognition,
        "use_chart_recognition": args.use_chart_recognition,
        "layout_batch_size": args.layout_batch_size,
        "text_det_batch_size": args.text_det_batch_size,
        "text_rec_batch_size": args.text_rec_batch_size,
    }

    for key, value in candidate_kwargs.items():
        if key in params and value is not None:
            kwargs[key] = value

    pipeline = PPStructureV3(**kwargs)
    return pipeline


def _save_result_json(res: Any, out_dir: Path, file_stem: str) -> Path:
    json_path = out_dir / f"{file_stem}.json"
    payload = getattr(res, "json", None)
    if callable(payload):
        payload = payload()
    if payload is None and hasattr(res, "res"):
        payload = getattr(res, "res")
        if callable(payload):
            payload = payload()
    if isinstance(payload, str):
        try:
            payload = json.loads(payload)
        except Exception:  # noqa: BLE001
            pass
    json_path.write_text(json.dumps(_json_ready(payload), ensure_ascii=False, indent=2), encoding="utf-8")
    return json_path


def _safe_save(callable_obj: Any, save_path: Path) -> str | None:
    try:
        callable_obj(str(save_path))
        return None
    except Exception as exc:  # noqa: BLE001
        return str(exc)


def _extract_markdown_info(res: Any) -> Any:
    markdown = getattr(res, "markdown", None)
    if callable(markdown):
        markdown = markdown()
    return markdown


def _write_combined_markdown(pipeline: Any, markdown_list: list[Any], out_path: Path) -> Path | None:
    if not markdown_list:
        return None
    if not hasattr(pipeline, "concatenate_markdown_pages"):
        return None
    merged = pipeline.concatenate_markdown_pages(markdown_list)
    if isinstance(merged, tuple):
        markdown_text = merged[0]
    elif isinstance(merged, dict):
        markdown_text = merged.get("markdown_texts") or merged.get("text") or ""
    else:
        markdown_text = merged
    if not isinstance(markdown_text, str):
        markdown_text = str(markdown_text)
    out_path.write_text(markdown_text, encoding="utf-8")
    return out_path


def process_document(
    pipeline: Any,
    input_path: Path,
    out_root: Path,
    args: argparse.Namespace,
) -> Path:
    doc_dir = _ensure_dir(out_root / _doc_id(input_path))
    json_dir = _ensure_dir(doc_dir / "json")
    md_dir = _ensure_dir(doc_dir / "markdown")
    html_dir = _ensure_dir(doc_dir / "html")
    xlsx_dir = _ensure_dir(doc_dir / "xlsx")
    img_dir = _ensure_dir(doc_dir / "images") if args.save_debug_images else None

    results = list(pipeline.predict(input=str(input_path)))
    if not results:
        raise RuntimeError(f"No PP-StructureV3 result returned for {input_path}")

    markdown_list: list[Any] = []
    summary_pages: list[dict[str, Any]] = []

    for page_idx, res in enumerate(results):
        page_name = f"{input_path.stem}_page_{page_idx + 1:03d}"
        page_summary: dict[str, Any] = {"page_index": page_idx}

        json_path = _save_result_json(res, json_dir, page_name)
        page_summary["json_path"] = str(json_path)

        if hasattr(res, "save_to_markdown"):
            err = _safe_save(res.save_to_markdown, md_dir)
            if err:
                page_summary["markdown_error"] = err
        markdown_info = _extract_markdown_info(res)
        if markdown_info is not None:
            markdown_list.append(markdown_info)
            md_info_path = md_dir / f"{page_name}.markdown_info.json"
            md_info_path.write_text(
                json.dumps(_json_ready(markdown_info), ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            page_summary["markdown_info_path"] = str(md_info_path)

        if hasattr(res, "save_to_html"):
            err = _safe_save(res.save_to_html, html_dir)
            if err:
                page_summary["html_error"] = err

        if hasattr(res, "save_to_xlsx"):
            err = _safe_save(res.save_to_xlsx, xlsx_dir)
            if err:
                page_summary["xlsx_error"] = err

        if img_dir is not None and hasattr(res, "save_to_img"):
            err = _safe_save(res.save_to_img, img_dir)
            if err:
                page_summary["image_error"] = err

        if hasattr(res, "json"):
            payload = getattr(res, "json")
            if callable(payload):
                payload = payload()
            page_summary["json_preview_keys"] = list((_json_ready(payload) or {}).keys()) if isinstance(payload, dict) else []

        summary_pages.append(page_summary)

    combined_md_path = _write_combined_markdown(
        pipeline=pipeline,
        markdown_list=markdown_list,
        out_path=doc_dir / f"{input_path.stem}.md",
    )

    summary = {
        "input_path": str(input_path),
        "page_count": len(results),
        "doc_dir": str(doc_dir),
        "combined_markdown_path": str(combined_md_path) if combined_md_path else None,
        "pages": summary_pages,
    }
    summary_path = doc_dir / "summary.json"
    summary_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    return summary_path


def build_arg_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(
        description="Parse PDFs/images with PP-StructureV3 and export structured JSON/Markdown/HTML/XLSX outputs."
    )
    ap.add_argument("--input_path", required=True, help="A PDF/image path or a directory containing them.")
    ap.add_argument(
        "--out_dir",
        default="out/pp_parse",
        help="Output directory for PP-StructureV3 artifacts.",
    )
    ap.add_argument("--recursive", action="store_true", help="Recursively scan the input directory.")
    ap.add_argument("--limit", type=int, default=0, help="Only process the first N documents.")
    ap.add_argument("--overwrite", action="store_true", help="Re-run even if summary.json already exists.")

    ap.add_argument("--device", default="", help="Inference device, e.g. gpu:0 or cpu.")
    ap.add_argument(
        "--disable_doc_orientation_classify",
        dest="use_doc_orientation_classify",
        action="store_false",
        help="Disable document orientation classification.",
    )
    ap.add_argument(
        "--disable_doc_unwarping",
        dest="use_doc_unwarping",
        action="store_false",
        help="Disable document unwarping.",
    )
    ap.add_argument(
        "--disable_table_recognition",
        dest="use_table_recognition",
        action="store_false",
        help="Disable table recognition.",
    )
    ap.add_argument(
        "--enable_seal_recognition",
        dest="use_seal_recognition",
        action="store_true",
        help="Enable seal recognition when supported.",
    )
    ap.add_argument(
        "--enable_formula_recognition",
        dest="use_formula_recognition",
        action="store_true",
        help="Enable formula recognition when supported.",
    )
    ap.add_argument(
        "--enable_chart_recognition",
        dest="use_chart_recognition",
        action="store_true",
        help="Enable chart recognition when supported.",
    )
    ap.add_argument("--layout_batch_size", type=int, default=None)
    ap.add_argument("--text_det_batch_size", type=int, default=None)
    ap.add_argument("--text_rec_batch_size", type=int, default=None)
    ap.add_argument("--save_debug_images", action="store_true", help="Save visualization images when supported.")

    ap.set_defaults(
        use_doc_orientation_classify=True,
        use_doc_unwarping=True,
        use_table_recognition=True,
        use_seal_recognition=False,
        use_formula_recognition=False,
        use_chart_recognition=False,
    )
    return ap


def main() -> int:
    parser = build_arg_parser()
    args = parser.parse_args()

    input_path = Path(args.input_path)
    out_dir = _ensure_dir(Path(args.out_dir))
    inputs = discover_inputs(input_path=input_path, recursive=bool(args.recursive))
    if args.limit > 0:
        inputs = inputs[: args.limit]
    if not inputs:
        raise FileNotFoundError(f"No PDF/image inputs found under: {input_path}")

    pipeline = init_pipeline(args)

    ok = 0
    skipped = 0
    failed = 0
    for idx, doc_path in enumerate(inputs, start=1):
        doc_dir = out_dir / _doc_id(doc_path)
        summary_path = doc_dir / "summary.json"
        if summary_path.exists() and not args.overwrite:
            skipped += 1
            print(f"[SKIP] {idx}/{len(inputs)} {doc_path}")
            continue
        try:
            out_summary = process_document(
                pipeline=pipeline,
                input_path=doc_path,
                out_root=out_dir,
                args=args,
            )
            ok += 1
            print(f"[OK] {idx}/{len(inputs)} {doc_path} -> {out_summary}")
        except Exception as exc:  # noqa: BLE001
            failed += 1
            _ensure_dir(doc_dir)
            (doc_dir / "error.txt").write_text(str(exc), encoding="utf-8")
            print(f"[ERR] {idx}/{len(inputs)} {doc_path} -> {exc}", file=sys.stderr)

    print(f"[DONE] total={len(inputs)} ok={ok} failed={failed} skipped={skipped} out_dir={out_dir}")
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
