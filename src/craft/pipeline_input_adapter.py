from __future__ import annotations

import importlib.util
import json
import shutil
import sys
from argparse import Namespace
from dataclasses import asdict, dataclass
from functools import lru_cache
from pathlib import Path
from types import ModuleType
from typing import Any, Callable


ROOT = Path(__file__).resolve().parent
CONVERT_DIR = ROOT / "convert"

WORD_SUFFIXES = {".docx"}
IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".bmp", ".webp", ".tif", ".tiff"}
PDF_SUFFIXES = {".pdf"}
EXCEL_PASSTHROUGH_SUFFIXES = {".xlsx"}
EXCEL_CONVERT_SUFFIXES = {".xls"}


@dataclass
class InputPlan:
    source_path: str
    source_kind: str
    needs_conversion: bool
    converter: str
    normalized_output_path: str


def _display_path(path: Path) -> str:
    try:
        return path.resolve().relative_to(ROOT.resolve()).as_posix()
    except Exception:
        return path.name


def _load_module(module_name: str, path: Path) -> ModuleType:
    spec = importlib.util.spec_from_file_location(module_name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Failed to load module from: {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module


@lru_cache(maxsize=None)
def _docx_module() -> ModuleType:
    return _load_module("docx_tables_runtime", CONVERT_DIR / "docx_tables.py")


@lru_cache(maxsize=None)
def _ppstructure_extract_module() -> ModuleType:
    return _load_module(
        "pp_parse_runtime",
        CONVERT_DIR / "pp_parse.py",
    )


@lru_cache(maxsize=None)
def _ppstructure_reconstruct_module() -> ModuleType:
    return _load_module(
        "pp_rebuild_runtime",
        CONVERT_DIR / "pp_rebuild.py",
    )


def _normalize_kind(src: Path) -> tuple[str, str]:
    suffix = src.suffix.lower()
    if suffix in EXCEL_PASSTHROUGH_SUFFIXES:
        return "excel", "passthrough"
    if suffix in EXCEL_CONVERT_SUFFIXES:
        return "excel", "xls_to_xlsx"
    if suffix in WORD_SUFFIXES:
        return "word", "docx_tables"
    if suffix in PDF_SUFFIXES:
        return "pdf", "pp_parse"
    if suffix in IMAGE_SUFFIXES:
        return "image", "pp_parse"
    raise ValueError(f"Unsupported input file type: {src}")


def plan_input_file(src: Path, normalized_output_path: Path) -> InputPlan:
    source_kind, converter = _normalize_kind(src)
    needs_conversion = converter != "passthrough"
    return InputPlan(
        source_path=_display_path(src),
        source_kind=source_kind,
        needs_conversion=needs_conversion,
        converter=converter,
        normalized_output_path=_display_path(normalized_output_path),
    )


def _convert_docx_to_xlsx(src: Path, dst: Path) -> None:
    module = _docx_module()
    tables = module.read_docx_tables(src)
    module.write_xlsx(tables, dst)


def _convert_pdf_or_image_to_xlsx(src: Path, dst: Path, work_dir: Path) -> None:
    extract_mod = _ppstructure_extract_module()
    reconstruct_mod = _ppstructure_reconstruct_module()

    extract_args = Namespace(
        device="",
        use_doc_orientation_classify=True,
        use_doc_unwarping=True,
        use_table_recognition=True,
        use_seal_recognition=False,
        use_formula_recognition=False,
        use_chart_recognition=False,
        layout_batch_size=None,
        text_det_batch_size=None,
        text_rec_batch_size=None,
        save_debug_images=False,
    )

    pp_out_dir = work_dir / "ppstructurev3"
    pp_out_dir.mkdir(parents=True, exist_ok=True)
    pipeline = extract_mod.init_pipeline(extract_args)
    summary_path = extract_mod.process_document(
        pipeline=pipeline,
        input_path=src,
        out_root=pp_out_dir,
        args=extract_args,
    )
    reconstructed_dir = work_dir / "reconstructed"
    reconstructed_dir.mkdir(parents=True, exist_ok=True)
    out_path = reconstruct_mod.convert_summary(
        summary_path=summary_path,
        out_dir=reconstructed_dir,
        column_width_mode="geometry",
    )
    shutil.copyfile(out_path, dst)


def ensure_pipeline_xlsx(
    src: Path,
    dst: Path,
    work_dir: Path,
    log_fn: Callable[[str], None],
    excel_ensure_fn: Callable[[Path, Path, Callable[[str], None]], Path],
) -> tuple[Path, InputPlan]:
    src = src.resolve()
    dst = dst.resolve()
    if not src.is_file():
        raise FileNotFoundError(f"Input file not found: {src}")

    plan = plan_input_file(src, dst)
    work_dir.mkdir(parents=True, exist_ok=True)

    log_fn(
        f"[INPUT_PLANNER] kind={plan.source_kind} needs_conversion={plan.needs_conversion} "
        f"converter={plan.converter} source={src.name}"
    )

    if plan.converter == "passthrough":
        if src != dst:
            shutil.copyfile(src, dst)
        return dst, plan

    if plan.converter == "xls_to_xlsx":
        return excel_ensure_fn(src, dst, log_fn), plan

    if plan.converter == "docx_tables":
        _convert_docx_to_xlsx(src, dst)
        return dst, plan

    if plan.converter == "pp_parse":
        _convert_pdf_or_image_to_xlsx(src, dst, work_dir=work_dir)
        return dst, plan

    raise ValueError(f"Unsupported conversion plan: {asdict(plan)}")


def write_plan_json(plan: InputPlan, path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(asdict(plan), ensure_ascii=False, indent=2), encoding="utf-8")
    return path
