from __future__ import annotations

import csv
import importlib.util
import json
import re
import sys
import zipfile
from argparse import Namespace
from dataclasses import asdict, dataclass, field
from email import policy
from email.parser import BytesParser
from functools import lru_cache
from html import unescape
from pathlib import Path
from typing import Any, Callable
from xml.etree import ElementTree as ET

import openpyxl

ROOT = Path(__file__).resolve().parent
CONVERT_DIR = ROOT / "convert"

TEXT_SUFFIXES = {".txt", ".md", ".rtf", ".log"}
STRUCTURED_SUFFIXES = {".json", ".xml", ".yaml", ".yml"}
TABLE_SUFFIXES = {".csv", ".tsv", ".xls", ".xlsx"}
HTML_SUFFIXES = {".html", ".htm", ".mhtml"}
EMAIL_SUFFIXES = {".eml"}
DOC_SUFFIXES = {".docx", ".doc", ".ppt", ".pptx"}
PDF_SUFFIXES = {".pdf"}
IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".bmp", ".webp", ".tif", ".tiff"}


@dataclass
class SourceAnalysis:
    path: str
    file_type: str
    format_family: str
    text_layer: str
    needs_ocr: bool
    semantic_structure: str
    layout_structure: str
    parse_mode: str
    representation: str
    routing_features: dict[str, Any] = field(default_factory=dict)
    fallback_modes: list[str] = field(default_factory=list)
    evidence_count: int = 0
    notes: list[str] = field(default_factory=list)


@dataclass
class SourceBundle:
    source_files: list[str]
    analyses: list[dict[str, Any]]
    normalized_documents: list[dict[str, Any]]
    synthesized_instruction: str
    original_instruction: str


def _display_path(path: Path) -> str:
    try:
        return path.resolve().relative_to(ROOT.resolve()).as_posix()
    except Exception:
        return path.name


def _load_module(module_name: str, path: Path) -> Any:
    spec = importlib.util.spec_from_file_location(module_name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Failed to load module from: {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module


@lru_cache(maxsize=None)
def _ppstructure_extract_module() -> Any:
    return _load_module(
        "pp_parse_source_runtime",
        CONVERT_DIR / "pp_parse.py",
    )


def _trim(text: str, limit: int = 240) -> str:
    clean = re.sub(r"\s+", " ", str(text or "")).strip()
    if len(clean) <= limit:
        return clean
    return clean[: max(0, limit - 3)] + "..."


def _read_text_file(path: Path) -> str:
    for enc in ("utf-8", "utf-8-sig", "gb18030", "latin-1"):
        try:
            return path.read_text(encoding=enc)
        except Exception:
            continue
    return path.read_text(errors="ignore")


def _guess_format_family(path: Path) -> tuple[str, str]:
    suffix = path.suffix.lower()
    if suffix in TEXT_SUFFIXES:
        return suffix.lstrip("."), "text"
    if suffix in STRUCTURED_SUFFIXES:
        return suffix.lstrip("."), "structured"
    if suffix in TABLE_SUFFIXES:
        return suffix.lstrip("."), "table"
    if suffix in HTML_SUFFIXES:
        return suffix.lstrip("."), "html"
    if suffix in EMAIL_SUFFIXES:
        return suffix.lstrip("."), "email"
    if suffix in DOC_SUFFIXES:
        return suffix.lstrip("."), "office"
    if suffix in PDF_SUFFIXES:
        return suffix.lstrip("."), "pdf"
    if suffix in IMAGE_SUFFIXES:
        return suffix.lstrip("."), "image"
    return suffix.lstrip(".") or "unknown", "unknown"


def _extract_json_like(path: Path) -> tuple[Any, list[dict[str, Any]], list[str]]:
    raw = _read_text_file(path)
    notes: list[str] = []
    suffix = path.suffix.lower()
    if suffix == ".json":
        return json.loads(raw), [], notes
    if suffix in {".yaml", ".yml"}:
        try:
            import yaml  # type: ignore

            return yaml.safe_load(raw), [], notes
        except Exception as exc:
            notes.append(f"yaml_parse_failed:{exc}")
            return {"raw_text": raw}, [], notes
    if suffix == ".xml":
        root = ET.fromstring(raw)
        return _xml_to_obj(root), [], notes
    return {"raw_text": raw}, [], notes


def _xml_to_obj(node: ET.Element) -> Any:
    children = list(node)
    if not children:
        text = (node.text or "").strip()
        if node.attrib:
            payload: dict[str, Any] = {"@attributes": dict(node.attrib)}
            if text:
                payload["#text"] = text
            return payload
        return text
    grouped: dict[str, list[Any]] = {}
    for child in children:
        grouped.setdefault(child.tag, []).append(_xml_to_obj(child))
    out: dict[str, Any] = {}
    if node.attrib:
        out["@attributes"] = dict(node.attrib)
    for key, items in grouped.items():
        out[key] = items[0] if len(items) == 1 else items
    text = (node.text or "").strip()
    if text:
        out["#text"] = text
    return out


def _extract_html(path: Path) -> tuple[dict[str, Any], list[dict[str, Any]], list[str]]:
    raw = _read_text_file(path)
    notes: list[str] = []
    tables: list[dict[str, Any]] = []
    text = ""
    try:
        from bs4 import BeautifulSoup  # type: ignore

        soup = BeautifulSoup(raw, "html.parser")
        text = soup.get_text("\n", strip=True)
        for idx, table in enumerate(soup.find_all("table"), start=1):
            rows: list[list[str]] = []
            for tr in table.find_all("tr"):
                cells = [cell.get_text(" ", strip=True) for cell in tr.find_all(["th", "td"])]
                if any(str(x).strip() for x in cells):
                    rows.append(cells)
            if rows:
                tables.append({"table_index": idx, "rows": rows[:50]})
    except Exception as exc:
        notes.append(f"bs4_unavailable_or_failed:{exc}")
        text = re.sub(r"<[^>]+>", " ", raw)
        text = unescape(text)
    return {"text": text, "tables": tables}, _text_evidence(text), notes


def _extract_tables_from_markdown(text: str) -> list[dict[str, Any]]:
    tables: list[dict[str, Any]] = []
    current: list[str] = []
    for line in str(text or "").splitlines():
        stripped = line.rstrip()
        if "|" in stripped:
            current.append(stripped)
            continue
        if len(current) >= 2:
            rows = _markdown_lines_to_rows(current)
            if rows:
                tables.append({"table_index": len(tables) + 1, "rows": rows[:50]})
        current = []
    if len(current) >= 2:
        rows = _markdown_lines_to_rows(current)
        if rows:
            tables.append({"table_index": len(tables) + 1, "rows": rows[:50]})
    return tables


def _markdown_lines_to_rows(lines: list[str]) -> list[list[str]]:
    rows: list[list[str]] = []
    for idx, line in enumerate(lines):
        stripped = line.strip().strip("|")
        cells = [cell.strip() for cell in stripped.split("|")]
        if idx == 1 and all(re.fullmatch(r"[:\- ]+", cell or "") for cell in cells):
            continue
        if any(cells):
            rows.append(cells)
    return rows


def _extract_tables_from_pp_page_json(value: Any, out: list[dict[str, Any]]) -> None:
    if isinstance(value, dict):
        rows = value.get("rows")
        if isinstance(rows, list) and rows and all(isinstance(row, list) for row in rows[:5]):
            out.append({"table_index": len(out) + 1, "rows": [[str(cell) for cell in row[:20]] for row in rows[:50]]})
        for item in value.values():
            _extract_tables_from_pp_page_json(item, out)
        return
    if isinstance(value, list):
        for item in value[:200]:
            _extract_tables_from_pp_page_json(item, out)


def _extract_with_ppstructure(path: Path, out_root: Path) -> tuple[dict[str, Any], list[dict[str, Any]], list[str]]:
    notes: list[str] = []
    text = ""
    tables: list[dict[str, Any]] = []
    summary_pages: list[dict[str, Any]] = []
    try:
        module = _ppstructure_extract_module()
        args = Namespace(
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
        out_root.mkdir(parents=True, exist_ok=True)
        pipeline = module.init_pipeline(args)
        summary_path = module.process_document(
            pipeline=pipeline,
            input_path=path,
            out_root=out_root,
            args=args,
        )
        summary = json.loads(Path(summary_path).read_text(encoding="utf-8"))
        summary_pages = list(summary.get("pages") or [])[:100]
        combined_md = str(summary.get("combined_markdown_path") or "").strip()
        if combined_md and Path(combined_md).exists():
            text = _read_text_file(Path(combined_md))
            tables.extend(_extract_tables_from_markdown(text))
        for page in summary_pages[:30]:
            json_path = str(page.get("json_path") or "").strip()
            if not json_path:
                continue
            page_json_path = Path(json_path)
            if not page_json_path.exists():
                continue
            try:
                payload = json.loads(page_json_path.read_text(encoding="utf-8"))
                _extract_tables_from_pp_page_json(payload, tables)
            except Exception as exc:
                notes.append(f"pp_page_json_parse_failed:{page_json_path.name}:{exc}")
        tables = tables[:20]
        evidence = _text_evidence(text, limit=20)
        for table in tables[:6]:
            rows = list(table.get("rows") or [])[:5]
            evidence.extend(_table_evidence(rows, sheet_name=f"pp_table_{table.get('table_index', '')}"))
        return {
            "text": text,
            "tables": tables,
            "pp_pages": summary_pages[:30],
            "ocr_engine": "ppstructurev3",
        }, evidence[:60], notes
    except Exception as exc:
        notes.append(f"ppstructure_ocr_failed:{exc}")
        return {"text": "", "tables": [], "ocr_engine": "ppstructurev3"}, [], notes


def _extract_email(path: Path) -> tuple[dict[str, Any], list[dict[str, Any]], list[str]]:
    with path.open("rb") as f:
        msg = BytesParser(policy=policy.default).parse(f)
    headers = {k: str(v) for k, v in msg.items()}
    body_parts: list[str] = []
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                try:
                    body_parts.append(part.get_content())
                except Exception:
                    continue
    else:
        try:
            body_parts.append(msg.get_content())
        except Exception:
            pass
    body = "\n".join(body_parts).strip()
    return {"headers": headers, "body": body}, _text_evidence(body), []


def _extract_csv(path: Path) -> tuple[dict[str, Any], list[dict[str, Any]], list[str]]:
    delimiter = "\t" if path.suffix.lower() == ".tsv" else ","
    rows: list[list[str]] = []
    with path.open("r", encoding="utf-8", errors="ignore", newline="") as f:
        reader = csv.reader(f, delimiter=delimiter)
        for idx, row in enumerate(reader):
            rows.append([str(x) for x in row])
            if idx >= 199:
                break
    return _tabular_representation(rows, sheet_name=path.name), _table_evidence(rows, sheet_name=path.name), []


def _extract_xlsx(path: Path) -> tuple[dict[str, Any], list[dict[str, Any]], list[str]]:
    notes: list[str] = []
    if path.suffix.lower() == ".xls":
        try:
            import pandas as pd  # type: ignore

            sheets: list[dict[str, Any]] = []
            evidence: list[dict[str, Any]] = []
            sheet_map = pd.read_excel(path, sheet_name=None, header=None)
            for name, frame in list(sheet_map.items())[:10]:
                rows = []
                for _, row in frame.head(80).iterrows():
                    rows.append(["" if value is None else str(value) for value in list(row)[:25]])
                rows = [row for row in rows if any(str(x).strip() for x in row)]
                if not rows:
                    continue
                sheets.append(_tabular_representation(rows, sheet_name=str(name)))
                evidence.extend(_table_evidence(rows[:20], sheet_name=str(name)))
            return {"sheets": sheets}, evidence[:60], notes
        except Exception as exc:
            notes.append(f"xls_read_failed:{exc}")
            return {"sheets": []}, [], notes
    wb = openpyxl.load_workbook(path, data_only=True)
    sheets: list[dict[str, Any]] = []
    evidence: list[dict[str, Any]] = []
    for ws in wb.worksheets[:10]:
        rows: list[list[str]] = []
        max_row = min(int(ws.max_row or 0), 80)
        max_col = min(int(ws.max_column or 0), 25)
        for r in range(1, max_row + 1):
            row_vals = []
            for c in range(1, max_col + 1):
                val = ws.cell(r, c).value
                row_vals.append("" if val is None else str(val))
            if any(str(x).strip() for x in row_vals):
                rows.append(row_vals)
        if not rows:
            continue
        sheet_payload = _tabular_representation(rows, sheet_name=ws.title)
        sheets.append(sheet_payload)
        evidence.extend(_table_evidence(rows[:20], sheet_name=ws.title))
    return {"sheets": sheets}, evidence[:60], notes


def _extract_docx(path: Path) -> tuple[dict[str, Any], list[dict[str, Any]], list[str]]:
    notes: list[str] = []
    paragraphs: list[str] = []
    tables: list[dict[str, Any]] = []
    try:
        import docx  # type: ignore

        doc = docx.Document(str(path))
        paragraphs = [p.text.strip() for p in doc.paragraphs if p.text and p.text.strip()]
        for idx, table in enumerate(doc.tables, start=1):
            rows = []
            for row in table.rows:
                cells = [cell.text.strip() for cell in row.cells]
                if any(cells):
                    rows.append(cells)
            if rows:
                tables.append({"table_index": idx, "rows": rows[:50]})
    except Exception as exc:
        notes.append(f"python_docx_unavailable_or_failed:{exc}")
        try:
            with zipfile.ZipFile(path) as zf:
                xml = zf.read("word/document.xml").decode("utf-8", errors="ignore")
            texts = re.findall(r"<w:t[^>]*>(.*?)</w:t>", xml)
            paragraphs = [unescape(t).strip() for t in texts if t.strip()]
        except Exception as zip_exc:
            notes.append(f"docx_zip_fallback_failed:{zip_exc}")
    text = "\n".join(paragraphs)
    return {"text": text, "tables": tables}, _text_evidence(text), notes


def _extract_pdf(path: Path, max_pages: int = 30) -> tuple[dict[str, Any], list[dict[str, Any]], list[str], bool]:
    notes: list[str] = []
    text_parts: list[str] = []
    try:
        try:
            from pypdf import PdfReader  # type: ignore
        except Exception:
            from PyPDF2 import PdfReader  # type: ignore
        reader = PdfReader(str(path))
        page_limit = max(0, int(max_pages or 0))
        pages = reader.pages if page_limit <= 0 else reader.pages[:page_limit]
        for idx, page in enumerate(pages, start=1):
            try:
                text = page.extract_text() or ""
            except Exception:
                text = ""
            if text.strip():
                text_parts.append(f"[Page {idx}]\n{text.strip()}")
    except Exception as exc:
        notes.append(f"pdf_text_extract_failed:{exc}")
    full_text = "\n\n".join(text_parts).strip()
    has_text = len(re.sub(r"\s+", "", full_text)) >= 40
    return {"text": full_text}, _text_evidence(full_text), notes, has_text


def _extract_image(path: Path) -> tuple[dict[str, Any], list[dict[str, Any]], list[str]]:
    notes: list[str] = []
    text = ""
    try:
        from PIL import Image  # type: ignore
        import pytesseract  # type: ignore

        text = pytesseract.image_to_string(Image.open(path))
    except Exception as exc:
        notes.append(f"ocr_unavailable_or_failed:{exc}")
    return {"text": text}, _text_evidence(text), notes


def _text_evidence(text: str, limit: int = 12) -> list[dict[str, Any]]:
    chunks = [c.strip() for c in re.split(r"\n{2,}", str(text or "")) if c.strip()]
    if not chunks:
        chunks = [c.strip() for c in str(text or "").splitlines() if c.strip()]
    out = []
    for idx, chunk in enumerate(chunks[:limit], start=1):
        out.append({"kind": "text_span", "index": idx, "text": _trim(chunk, 220)})
    return out


def _table_evidence(rows: list[list[str]], sheet_name: str = "") -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    for idx, row in enumerate(rows[:10], start=1):
        out.append(
            {
                "kind": "table_row",
                "sheet": sheet_name,
                "row_index": idx,
                "values": [_trim(x, 80) for x in row[:12]],
            }
        )
    return out


def _tabular_representation(rows: list[list[str]], sheet_name: str = "") -> dict[str, Any]:
    non_empty_rows = [row for row in rows if any(str(x).strip() for x in row)]
    header = non_empty_rows[0] if non_empty_rows else []
    looks_header = len(set([str(x).strip().lower() for x in header if str(x).strip()])) >= max(1, len([x for x in header if str(x).strip()]) - 1)
    records: list[dict[str, Any]] = []
    if looks_header and len(non_empty_rows) > 1:
        keys = [str(x).strip() or f"col_{idx+1}" for idx, x in enumerate(header)]
        for row in non_empty_rows[1:21]:
            record = {keys[idx]: (row[idx] if idx < len(row) else "") for idx in range(len(keys))}
            if any(str(v).strip() for v in record.values()):
                records.append(record)
    return {
        "sheet_name": sheet_name,
        "header": header[:20],
        "row_count_preview": len(non_empty_rows[:20]),
        "column_count_preview": max((len(row) for row in non_empty_rows[:20]), default=0),
        "rows_preview": [row[:20] for row in non_empty_rows[:20]],
        "records_preview": records[:20],
    }


def _infer_text_semantics(text: str) -> tuple[str, str, str]:
    stripped = str(text or "").strip()
    if not stripped:
        return "sparse_or_unreadable", "mixed_or_unknown", "evidence_grounded_spans"
    lines = [ln.strip() for ln in stripped.splitlines() if ln.strip()]
    section_hits = sum(1 for ln in lines if re.match(r"^([A-Z][A-Za-z0-9 /&()-]{1,80}:|#+\s+.+)$", ln))
    kv_hits = sum(1 for ln in lines if ":" in ln and len(ln.split(":", 1)[0].strip()) <= 60)
    bullet_hits = sum(1 for ln in lines if re.match(r"^([-*]|\d+\.)\s+", ln))
    if kv_hits >= max(3, len(lines) // 4):
        return "flat_attributes_or_local_kv", "linear_text", "flat_kv_pairs"
    if section_hits >= 2:
        return "sectioned_document", "block_document", "sectioned_json"
    if bullet_hits >= 3:
        return "record_collection", "block_document", "list_of_records"
    return "free_text_document", "linear_text", "evidence_grounded_spans"


def _infer_object_semantics(obj: Any) -> tuple[str, str]:
    if isinstance(obj, dict):
        return "hierarchical_business_document", "hierarchical_json"
    if isinstance(obj, list):
        if obj and all(isinstance(item, dict) for item in obj[:5]):
            return "record_collection", "list_of_records"
        return "repeated_items", "list_of_records"
    return "flat_attributes_or_local_kv", "flat_kv_pairs"


def _normalize_object(obj: Any, depth: int = 0) -> Any:
    if depth >= 5:
        return _trim(str(obj), 400)
    if isinstance(obj, dict):
        out: dict[str, Any] = {}
        for idx, (k, v) in enumerate(obj.items()):
            if idx >= 60:
                out["__truncated__"] = True
                break
            out[str(k)] = _normalize_object(v, depth + 1)
        return out
    if isinstance(obj, list):
        return [_normalize_object(v, depth + 1) for v in obj[:30]]
    if obj is None:
        return None
    return _trim(str(obj), 400)


def _extract_flat_kv_from_text(text: str) -> dict[str, Any]:
    pairs: list[dict[str, str]] = []
    for idx, line in enumerate(str(text or "").splitlines(), start=1):
        stripped = line.strip()
        if not stripped or ":" not in stripped:
            continue
        key, value = stripped.split(":", 1)
        key = key.strip()
        value = value.strip()
        if not key or len(key) > 80:
            continue
        pairs.append({"key": key, "value": value, "line_index": idx})
        if len(pairs) >= 60:
            break
    return {"pairs": pairs}


def _extract_sectioned_from_text(text: str) -> dict[str, Any]:
    sections: list[dict[str, Any]] = []
    current = {"title": "root", "content": []}
    for idx, line in enumerate(str(text or "").splitlines(), start=1):
        stripped = line.strip()
        if not stripped:
            continue
        if re.match(r"^(#+\s+.+|[A-Z][A-Za-z0-9 /&()-]{1,80}:)$", stripped):
            if current["content"] or current["title"] != "root":
                sections.append(current)
            current = {"title": stripped.rstrip(":"), "content": []}
        else:
            current["content"].append({"line_index": idx, "text": stripped})
    if current["content"] or current["title"] != "root":
        sections.append(current)
    return {"sections": sections[:40]}


def _extract_records_from_text(text: str) -> dict[str, Any]:
    records: list[dict[str, Any]] = []
    current: list[str] = []
    for line in str(text or "").splitlines():
        stripped = line.strip()
        if not stripped:
            if current:
                records.append({"items": current[:20]})
                current = []
            continue
        if re.match(r"^([-*]|\d+\.)\s+", stripped):
            if current and len(current) >= 3:
                records.append({"items": current[:20]})
                current = []
            current.append(re.sub(r"^([-*]|\d+\.)\s+", "", stripped))
        else:
            current.append(stripped)
    if current:
        records.append({"items": current[:20]})
    return {"records": records[:30]}


def _extract_hierarchical_from_text(text: str) -> dict[str, Any]:
    root: list[dict[str, Any]] = []
    stack: list[tuple[int, list[dict[str, Any]]]] = [(0, root)]
    for idx, line in enumerate(str(text or "").splitlines(), start=1):
        if not line.strip():
            continue
        indent = len(line) - len(line.lstrip(" "))
        node = {"text": line.strip(), "line_index": idx, "children": []}
        while len(stack) > 1 and indent < stack[-1][0]:
            stack.pop()
        stack[-1][1].append(node)
        stack.append((indent + 1, node["children"]))
    return {"nodes": root[:50]}


def _extract_table_like(normalized: Any) -> dict[str, Any]:
    if isinstance(normalized, dict):
        if isinstance(normalized.get("sheets"), list):
            return {"tables": normalized.get("sheets")[:10]}
        if isinstance(normalized.get("rows_preview"), list):
            return {
                "tables": [
                    {
                        "sheet_name": normalized.get("sheet_name", ""),
                        "header": normalized.get("header", []),
                        "rows_preview": normalized.get("rows_preview", []),
                        "records_preview": normalized.get("records_preview", []),
                    }
                ]
            }
        if isinstance(normalized.get("tables"), list):
            return {"tables": normalized.get("tables")[:10], "text": normalized.get("text", "")}
    return {"tables": []}


def _extract_list_of_records(normalized: Any) -> dict[str, Any]:
    records: list[Any] = []
    if isinstance(normalized, dict):
        if isinstance(normalized.get("records_preview"), list):
            records.extend(normalized.get("records_preview") or [])
        if isinstance(normalized.get("sheets"), list):
            for sheet in normalized.get("sheets") or []:
                if isinstance(sheet, dict) and isinstance(sheet.get("records_preview"), list):
                    for item in sheet.get("records_preview") or []:
                        records.append({"sheet_name": sheet.get("sheet_name"), "record": item})
        if isinstance(normalized.get("tables"), list):
            for table in normalized.get("tables") or []:
                if isinstance(table, dict) and isinstance(table.get("rows"), list):
                    rows = table.get("rows") or []
                    if not rows:
                        continue
                    header = [str(x).strip() or f"col_{i+1}" for i, x in enumerate(rows[0])]
                    for row in rows[1:21]:
                        records.append({header[i]: row[i] if i < len(row) else "" for i in range(len(header))})
    return {"records": records[:40]}


def _extract_hybrid_sections_and_tables(normalized: Any) -> dict[str, Any]:
    text = ""
    tables: list[Any] = []
    if isinstance(normalized, dict):
        text = str(normalized.get("text") or "")
        if isinstance(normalized.get("tables"), list):
            tables = normalized.get("tables")[:10]
        if isinstance(normalized.get("sheets"), list):
            tables = normalized.get("sheets")[:10]
    return {
        "sections": _extract_sectioned_from_text(text).get("sections", []) if text else [],
        "tables": tables,
    }


def _extract_evidence_spans(text: str, evidence: list[dict[str, Any]]) -> dict[str, Any]:
    return {
        "text_preview": _trim(text, 4000),
        "evidence_spans": evidence[:40],
    }


def _route_specific_structuring(
    *,
    raw_material: Any,
    representation: str,
    evidence: list[dict[str, Any]],
) -> Any:
    text = _flatten_text_from_normalized(raw_material)
    if representation == "flat_kv_pairs":
        return _extract_flat_kv_from_text(text)
    if representation == "sectioned_json":
        return _extract_sectioned_from_text(text)
    if representation == "hierarchical_json":
        if isinstance(raw_material, (dict, list)):
            return _normalize_object(raw_material)
        return _extract_hierarchical_from_text(text)
    if representation == "list_of_records":
        structured = _extract_list_of_records(raw_material)
        if structured.get("records"):
            return structured
        return _extract_records_from_text(text)
    if representation == "table_like_representation":
        return _extract_table_like(raw_material)
    if representation == "hybrid_sectioned_json_and_table_blocks":
        return _extract_hybrid_sections_and_tables(raw_material)
    return _extract_evidence_spans(text, evidence)


def _flatten_text_from_normalized(normalized: Any) -> str:
    chunks: list[str] = []

    def visit(value: Any, depth: int = 0) -> None:
        if depth > 4:
            return
        if isinstance(value, dict):
            for key, item in value.items():
                if key in {"evidence", "analysis"}:
                    continue
                visit(item, depth + 1)
            return
        if isinstance(value, list):
            for item in value[:50]:
                visit(item, depth + 1)
            return
        text = str(value or "").strip()
        if text:
            chunks.append(text)

    visit(normalized)
    return "\n".join(chunks)


def _object_depth(value: Any, depth: int = 0) -> int:
    if isinstance(value, dict):
        if not value:
            return depth
        return max(_object_depth(v, depth + 1) for v in value.values())
    if isinstance(value, list):
        if not value:
            return depth
        return max(_object_depth(v, depth + 1) for v in value[:20])
    return depth


def _estimate_record_count(normalized: Any) -> int:
    if isinstance(normalized, dict):
        records = normalized.get("records_preview")
        if isinstance(records, list):
            return len(records)
        sheets = normalized.get("sheets")
        if isinstance(sheets, list):
            total = 0
            for sheet in sheets[:10]:
                if isinstance(sheet, dict) and isinstance(sheet.get("records_preview"), list):
                    total += len(sheet.get("records_preview") or [])
            return total
    if isinstance(normalized, list):
        return len(normalized)
    return 0


def _estimate_table_shape(normalized: Any) -> tuple[int, int, int]:
    row_count = 0
    col_count = 0
    table_blocks = 0
    if isinstance(normalized, dict):
        if isinstance(normalized.get("rows_preview"), list):
            rows = normalized.get("rows_preview") or []
            table_blocks += 1
            row_count = max(row_count, len(rows))
            col_count = max(col_count, max((len(row) for row in rows if isinstance(row, list)), default=0))
        if isinstance(normalized.get("tables"), list):
            for table in normalized.get("tables") or []:
                if not isinstance(table, dict):
                    continue
                rows = table.get("rows") or []
                if isinstance(rows, list):
                    table_blocks += 1
                    row_count = max(row_count, len(rows))
                    col_count = max(col_count, max((len(row) for row in rows if isinstance(row, list)), default=0))
        if isinstance(normalized.get("sheets"), list):
            for sheet in normalized.get("sheets") or []:
                if not isinstance(sheet, dict):
                    continue
                rows = sheet.get("rows_preview") or []
                if isinstance(rows, list):
                    table_blocks += 1
                    row_count = max(row_count, len(rows))
                    col_count = max(col_count, max((len(row) for row in rows if isinstance(row, list)), default=0))
    return row_count, col_count, table_blocks


def _compute_routing_features(normalized: Any, evidence: list[dict[str, Any]]) -> dict[str, Any]:
    text = _flatten_text_from_normalized(normalized)
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    line_count = len(lines)
    char_count = len(text)
    kv_hits = sum(1 for ln in lines if ":" in ln and len(ln.split(":", 1)[0].strip()) <= 60)
    section_hits = sum(1 for ln in lines if re.match(r"^([A-Z][A-Za-z0-9 /&()-]{1,80}:|#+\s+.+)$", ln))
    bullet_hits = sum(1 for ln in lines if re.match(r"^([-*]|\d+\.)\s+", ln))
    indent_hits = sum(1 for ln in lines if re.match(r"^\s{2,}\S+", ln))
    table_row_count, table_col_count, table_blocks = _estimate_table_shape(normalized)
    record_count = _estimate_record_count(normalized)
    object_depth = _object_depth(normalized)
    non_empty_density = 0.0
    if table_row_count and table_col_count and isinstance(normalized, dict):
        rows = normalized.get("rows_preview")
        if not isinstance(rows, list):
            sheets = normalized.get("sheets") or []
            if isinstance(sheets, list) and sheets and isinstance(sheets[0], dict):
                rows = sheets[0].get("rows_preview")
        if isinstance(rows, list) and rows:
            total_cells = sum(len(row) for row in rows if isinstance(row, list))
            filled_cells = sum(
                1
                for row in rows if isinstance(row, list)
                for cell in row
                if str(cell or "").strip()
            )
            if total_cells > 0:
                non_empty_density = round(filled_cells / total_cells, 4)
    features = {
        "char_count": char_count,
        "line_count": line_count,
        "content_size": (
            "small" if char_count < 600 else
            "medium" if char_count < 4000 else
            "large"
        ),
        "kv_line_ratio": round(kv_hits / max(1, line_count), 4),
        "section_line_ratio": round(section_hits / max(1, line_count), 4),
        "bullet_line_ratio": round(bullet_hits / max(1, line_count), 4),
        "indent_ratio": round(indent_hits / max(1, line_count), 4),
        "object_depth": object_depth,
        "hierarchy_strength": (
            "high" if object_depth >= 4 or indent_hits >= 4 else
            "medium" if object_depth >= 2 or section_hits >= 2 else
            "low"
        ),
        "record_count": record_count,
        "repetition_strength": (
            "high" if record_count >= 8 or bullet_hits >= 5 else
            "medium" if record_count >= 2 or bullet_hits >= 2 else
            "low"
        ),
        "table_blocks": table_blocks,
        "table_row_count": table_row_count,
        "table_col_count": table_col_count,
        "grid_strength": (
            "high" if table_blocks >= 1 and table_row_count >= 4 and table_col_count >= 3 else
            "medium" if table_blocks >= 1 and table_col_count >= 2 else
            "low"
        ),
        "non_empty_density": non_empty_density,
        "sparsity": (
            "high" if 0 < non_empty_density <= 0.25 else
            "medium" if non_empty_density <= 0.6 else
            "low"
        ),
        "evidence_count": len(evidence),
    }
    return features


def _route_from_features(
    *,
    family: str,
    file_type: str,
    normalized: Any,
    evidence: list[dict[str, Any]],
    text_layer: str,
    needs_ocr: bool,
) -> tuple[str, str, str, list[str], dict[str, Any]]:
    features = _compute_routing_features(normalized, evidence)
    fallback_modes: list[str] = []
    semantic_structure = "free_text_document"
    layout_structure = "linear_text"
    parse_mode = "direct_text_parsing"
    representation = "evidence_grounded_spans"

    if needs_ocr:
        parse_mode = "ocr"
        fallback_modes.extend(["layout_parsing", "visual_understanding"])
        if features["grid_strength"] in {"medium", "high"}:
            layout_structure = "table_layout"
            semantic_structure = "tabular_or_repeated_records"
            representation = "table_like_representation"
            parse_mode = "ocr_plus_layout_parsing"
        else:
            layout_structure = "mixed_or_unknown"
            semantic_structure = "sparse_or_unreadable"
            representation = "evidence_grounded_spans"
        return semantic_structure, layout_structure, parse_mode, representation, fallback_modes, features

    if family == "structured":
        parse_mode = "structured_data_parsing"
    elif family == "table":
        parse_mode = "table_extraction"
    elif family == "html":
        parse_mode = "html_dom_parsing"
    elif family == "email":
        parse_mode = "email_parsing"
    elif family == "office":
        parse_mode = "office_document_parsing"
    elif family == "pdf" and text_layer == "native":
        parse_mode = "pdf_text_layer_parsing"

    if features["grid_strength"] == "high":
        layout_structure = "table_layout"
        semantic_structure = "tabular_or_repeated_records"
        representation = "table_like_representation"
        if features["hierarchy_strength"] in {"medium", "high"}:
            representation = "hybrid_sectioned_json_and_table_blocks"
            layout_structure = "mixed_blocks_and_tables"
        if features["repetition_strength"] == "high" and features["table_col_count"] <= 8:
            representation = "list_of_records"
    elif features["hierarchy_strength"] == "high":
        semantic_structure = "hierarchical_business_document"
        layout_structure = "hierarchical_document"
        representation = "hierarchical_json"
    elif features["section_line_ratio"] >= 0.12 or features["hierarchy_strength"] == "medium":
        semantic_structure = "sectioned_document"
        layout_structure = "block_document"
        representation = "sectioned_json"
    elif features["repetition_strength"] == "high":
        semantic_structure = "record_collection"
        layout_structure = "block_document"
        representation = "list_of_records"
    elif features["kv_line_ratio"] >= 0.28 and features["content_size"] in {"small", "medium"}:
        semantic_structure = "flat_attributes_or_local_kv"
        layout_structure = "linear_text"
        representation = "flat_kv_pairs"
    elif features["content_size"] == "small" and features["kv_line_ratio"] >= 0.15:
        semantic_structure = "flat_attributes_or_local_kv"
        layout_structure = "linear_text"
        representation = "flat_kv_pairs"
    else:
        semantic_structure = "free_text_document"
        layout_structure = "linear_text" if features["grid_strength"] == "low" else "mixed_blocks_and_tables"
        representation = "evidence_grounded_spans"

    if family in {"pdf", "image"} and features["grid_strength"] in {"medium", "high"}:
        fallback_modes.append("layout_parsing")
    if family in {"html", "pdf", "image"} and representation == "evidence_grounded_spans":
        fallback_modes.append("visual_understanding")
    if file_type in {"doc", "ppt", "pptx"}:
        fallback_modes.extend(["manual_conversion", "layout_parsing"])
    fallback_modes = list(dict.fromkeys([mode for mode in fallback_modes if mode]))
    return semantic_structure, layout_structure, parse_mode, representation, fallback_modes, features


def analyze_and_structure_source(
    path: Path,
    *,
    pdf_max_pages: int = 30,
    ocr_work_dir: Path | None = None,
) -> tuple[SourceAnalysis, dict[str, Any]]:
    file_type, family = _guess_format_family(path)
    notes: list[str] = []
    raw_material: Any
    evidence: list[dict[str, Any]]
    text_layer = "unknown"
    needs_ocr = False
    semantic_structure = "unknown"
    layout_structure = "mixed_or_unknown"
    parse_mode = "direct_text_parsing"
    representation = "evidence_grounded_spans"
    routing_features: dict[str, Any] = {}
    fallback_modes: list[str] = []

    if family == "text":
        text = _read_text_file(path)
        raw_material = {"text": text}
        evidence = _text_evidence(text)
        text_layer = "native"
    elif family == "structured":
        obj, evidence, extra_notes = _extract_json_like(path)
        notes.extend(extra_notes)
        raw_material = obj
        text_layer = "native"
    elif family == "table":
        if file_type in {"csv", "tsv"}:
            raw_material, evidence, extra_notes = _extract_csv(path)
        else:
            raw_material, evidence, extra_notes = _extract_xlsx(path)
        notes.extend(extra_notes)
        text_layer = "native"
    elif family == "html":
        raw_material, evidence, extra_notes = _extract_html(path)
        notes.extend(extra_notes)
        text_layer = "native"
    elif family == "email":
        raw_material, evidence, extra_notes = _extract_email(path)
        notes.extend(extra_notes)
        text_layer = "native"
    elif family == "office":
        if file_type == "docx":
            raw_material, evidence, extra_notes = _extract_docx(path)
            notes.extend(extra_notes)
            text_layer = "native"
        else:
            raw_material = {"raw_hint": f"Unsupported office format for direct parsing: {file_type}"}
            evidence = []
            notes.append("office_format_not_yet_supported_for_direct_parsing")
            text_layer = "unknown"
    elif family == "pdf":
        raw_material, evidence, extra_notes, has_text = _extract_pdf(path, max_pages=pdf_max_pages)
        notes.extend(extra_notes)
        if has_text:
            text_layer = "native"
        else:
            text_layer = "missing"
            needs_ocr = True
            if ocr_work_dir is not None:
                ocr_material, ocr_evidence, ocr_notes = _extract_with_ppstructure(path, ocr_work_dir / path.stem)
                if (ocr_material.get("text") or "").strip() or (ocr_material.get("tables") or []):
                    raw_material = ocr_material
                    evidence = ocr_evidence
                notes.extend(ocr_notes)
    elif family == "image":
        if ocr_work_dir is not None:
            raw_material, evidence, extra_notes = _extract_with_ppstructure(path, ocr_work_dir / path.stem)
            notes.extend(extra_notes)
            if not (raw_material.get("text") or "").strip() and not (raw_material.get("tables") or []):
                raw_material, evidence, extra_notes = _extract_image(path)
                notes.extend(extra_notes)
        else:
            raw_material, evidence, extra_notes = _extract_image(path)
            notes.extend(extra_notes)
        text_layer = "missing"
        needs_ocr = True
    else:
        raw_material = {"raw_text": _read_text_file(path)}
        evidence = _text_evidence(str(raw_material.get("raw_text") or ""))
        text_layer = "native"

    semantic_structure, layout_structure, parse_mode, representation, fallback_modes, routing_features = _route_from_features(
        family=family,
        file_type=file_type,
        normalized=raw_material,
        evidence=evidence,
        text_layer=text_layer,
        needs_ocr=needs_ocr,
    )
    normalized_representation = _route_specific_structuring(
        raw_material=raw_material,
        representation=representation,
        evidence=evidence,
    )

    analysis = SourceAnalysis(
        path=_display_path(path),
        file_type=file_type,
        format_family=family,
        text_layer=text_layer,
        needs_ocr=needs_ocr,
        semantic_structure=semantic_structure,
        layout_structure=layout_structure,
        routing_features=routing_features,
        parse_mode=parse_mode,
        representation=representation,
        fallback_modes=fallback_modes,
        evidence_count=len(evidence),
        notes=notes,
    )
    structured = {
        "source_path": _display_path(path),
        "analysis": asdict(analysis),
        "raw_material": _normalize_object(raw_material),
        "normalized_representation": normalized_representation,
        "evidence": evidence[:40],
    }
    return analysis, structured


def _load_source_paths(source_files: list[str] | None, manifest_path: str | None) -> list[Path]:
    items: list[str] = []
    for item in source_files or []:
        text = str(item or "").strip()
        if text:
            items.append(text)
    if manifest_path:
        manifest = Path(manifest_path).resolve()
        raw = _read_text_file(manifest)
        try:
            payload = json.loads(raw)
            if isinstance(payload, list):
                items.extend([str(x) for x in payload if str(x).strip()])
            elif isinstance(payload, dict):
                for key in ("files", "paths", "source_files"):
                    vals = payload.get(key)
                    if isinstance(vals, list):
                        items.extend([str(x) for x in vals if str(x).strip()])
                        break
        except Exception:
            for line in raw.splitlines():
                line = line.strip()
                if line:
                    items.append(line)
    deduped: list[Path] = []
    seen: set[str] = set()
    for item in items:
        path = Path(item)
        if not path.is_absolute():
            path = (Path.cwd() / path).resolve()
        else:
            path = path.resolve()
        key = str(path).lower()
        if key in seen:
            continue
        seen.add(key)
        deduped.append(path)
    return deduped


def build_source_bundle(
    source_files: list[str] | None,
    source_manifest_path: str | None,
    original_instruction: str,
    out_dir: Path,
    log_fn: Callable[[str], None],
    max_inline_chars: int = 24000,
    pdf_max_pages: int = 30,
) -> SourceBundle | None:
    paths = _load_source_paths(source_files, source_manifest_path)
    if not paths:
        return None
    out_dir.mkdir(parents=True, exist_ok=True)
    analyses: list[dict[str, Any]] = []
    documents: list[dict[str, Any]] = []
    for path in paths:
        if not path.exists():
            raise FileNotFoundError(f"Source file not found: {path}")
        analysis, structured = analyze_and_structure_source(
            path,
            pdf_max_pages=pdf_max_pages,
            ocr_work_dir=out_dir / "ppstructurev3",
        )
        analyses.append(asdict(analysis))
        documents.append(structured)
        log_fn(
            f"[SOURCE_ADAPTER] file={path.name} family={analysis.format_family} "
            f"parse_mode={analysis.parse_mode} repr={analysis.representation}"
        )
    payload = {
        "source_files": [str(p) for p in paths],
        "analyses": analyses,
        "normalized_documents": documents,
        "original_instruction": original_instruction,
    }
    json_path = out_dir / "source_bundle.json"
    json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    synthesized_instruction = _build_synthesized_instruction(
        original_instruction=original_instruction,
        analyses=analyses,
        documents=documents,
        max_inline_chars=max_inline_chars,
    )
    (out_dir / "source_instruction.txt").write_text(synthesized_instruction, encoding="utf-8")
    return SourceBundle(
        source_files=[_display_path(p) for p in paths],
        analyses=analyses,
        normalized_documents=documents,
        synthesized_instruction=synthesized_instruction,
        original_instruction=original_instruction,
    )


def _build_synthesized_instruction(
    original_instruction: str,
    analyses: list[dict[str, Any]],
    documents: list[dict[str, Any]],
    max_inline_chars: int,
) -> str:
    parts: list[str] = []
    base = str(original_instruction or "").strip()
    if base:
        parts.append("Primary instruction:")
        parts.append(base)
        parts.append("")
    parts.append("Adaptive source context:")
    for idx, analysis in enumerate(analyses, start=1):
        parts.append(
            f"- Source {idx}: path={analysis.get('path')} "
            f"type={analysis.get('file_type')} family={analysis.get('format_family')} "
            f"parse_mode={analysis.get('parse_mode')} representation={analysis.get('representation')} "
            f"semantic={analysis.get('semantic_structure')} layout={analysis.get('layout_structure')} "
            f"text_layer={analysis.get('text_layer')} needs_ocr={analysis.get('needs_ocr')}"
        )
        routing_features = analysis.get("routing_features") or {}
        if routing_features:
            parts.append(f"  routing_features={json.dumps(routing_features, ensure_ascii=False)}")
        notes = analysis.get("notes") or []
        if notes:
            parts.append(f"  notes={json.dumps(notes, ensure_ascii=False)}")
    parts.append("")
    parts.append("Structured source evidence:")
    compact_docs = []
    for doc in documents:
        compact_docs.append(
            {
                "source_path": doc.get("source_path"),
                "analysis": doc.get("analysis"),
                "raw_material": doc.get("raw_material"),
                "normalized_representation": doc.get("normalized_representation"),
                "evidence": doc.get("evidence"),
            }
        )
    doc_json = json.dumps(compact_docs, ensure_ascii=False, indent=2)
    if max_inline_chars > 0 and len(doc_json) > max_inline_chars:
        doc_json = doc_json[:max_inline_chars] + "\n...<truncated>"
    parts.append(doc_json)
    parts.append("")
    parts.append(
        "Use the adaptive source context as the grounding source for field values. "
        "The raw_material block is the broad extracted substrate; normalized_representation is the route-specific structured view chosen by the planner. "
        "Prefer explicit evidence over guessing. Preserve blanks when a field is not supported by the sources."
    )
    return "\n".join(parts).strip() + "\n"
