from __future__ import annotations

import argparse
import json
import sys
from dataclasses import asdict
from pathlib import Path
from typing import Any

try:
    from .pipeline_agent_actions import (
        build_runtime_paths,
        create_usage_summary,
        finalize_usage_summary,
        free_fill,
        guided_refill,
        load_initial_form,
        load_planner_policy,
        load_skills_text,
        rebuild_from_verification,
        reflect_and_hint,
        restore_template_region,
        verify_and_deliver,
    )
    from .pipeline_agent_planner import build_bootstrap_plan, plan_next_action
    from .pipeline_agent_types import ControllerState, ReflectRegionPlan
    from .pipeline_source_adapter import build_source_bundle
    from . import my_pipeline_legacy as legacy
except ImportError:
    from pipeline_agent_actions import (  # type: ignore
        build_runtime_paths,
        create_usage_summary,
        finalize_usage_summary,
        free_fill,
        guided_refill,
        load_initial_form,
        load_planner_policy,
        load_skills_text,
        rebuild_from_verification,
        reflect_and_hint,
        restore_template_region,
        verify_and_deliver,
    )
    from pipeline_agent_planner import build_bootstrap_plan, plan_next_action  # type: ignore
    from pipeline_agent_types import ControllerState, ReflectRegionPlan  # type: ignore
    from pipeline_source_adapter import build_source_bundle  # type: ignore
    import my_pipeline_legacy as legacy  # type: ignore


PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_SEGFORMER_CKPT = PROJECT_ROOT / "src" / "craft" / "vision" / "best_segformer.pt"
DEFAULT_AGENT_SKILLS_DIR = PROJECT_ROOT / "agent_skills"
DEFAULT_CONFIG_PATH = PROJECT_ROOT / "config.yaml"


def _log(msg: str) -> None:
    legacy._emit_console_line(msg)


def _write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding="utf-8")


def _find_region_plan(state: ControllerState, region: str) -> ReflectRegionPlan:
    if state.current_reflect is None:
        raise legacy.PipelineError("No active reflect stage is available.")
    for item in state.current_reflect.regions:
        if item.region == region:
            return item
    raise legacy.PipelineError(f"Region not found in current reflect plan: {region}")


def _record_action(state: ControllerState, payload: dict[str, Any]) -> None:
    state.action_history.append(payload)
    _write_json(state.paths.controller_debug_dir / "action_history.json", state.action_history)


def _build_config_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(add_help=False)
    ap.add_argument("--config", default=str(DEFAULT_CONFIG_PATH), help="Path to YAML config file.")
    return ap


def _flatten_config(raw: Any) -> dict[str, Any]:
    if not isinstance(raw, dict):
        return {}
    out: dict[str, Any] = {}
    valid_top_level = {
        "planner_json_path",
        "agentic_planner_enable",
        "planner_model",
        "planner_base_url",
        "planner_api_key",
        "planner_temperature",
        "planner_debug_dir",
        "save_model_io",
        "model_io_dump_dir",
        "runtime_job_dir",
        "runtime_root_dir",
        "dpi",
        "ckpt",
        "preprocess_backend",
        "render_canvas_bg",
        "render_png_scale",
        "render_max_rows",
        "render_max_cols",
        "agent_model",
        "agent_base_url",
        "agent_api_key",
        "agent_temperature",
        "agent_form_context_mode",
        "agent_include_rendered_html",
        "agent_rendered_html_max_chars",
        "tool_mode",
        "max_steps",
        "external_qwen_pairs_json",
        "agent_skills",
        "agent_skills_file",
        "reflect_enable",
        "reflect_plugin_enable",
        "reflect_max_rounds",
        "reflect_patch_steps",
        "reflect_checker_model",
        "reflect_edit_log_jsonl",
        "reflect_risk_threshold",
        "reflect_expand_left_cols",
        "reflect_expand_right_cols",
        "reflect_expand_top_rows",
        "reflect_expand_bottom_rows",
        "reflect_focus_pad_left_cols",
        "reflect_focus_pad_top_rows",
        "reflect_focus_pad_right_cols",
        "reflect_focus_pad_bottom_rows",
        "controller_max_cycles",
        "source_inline_max_chars",
        "source_pdf_max_pages",
        "source_manifest",
        "instruction",
        "sheet",
    }
    section_mappings = {
        "planner": {
            "enable": "agentic_planner_enable",
            "json_path": "planner_json_path",
            "model": "planner_model",
            "base_url": "planner_base_url",
            "api_key": "planner_api_key",
            "temperature": "planner_temperature",
            "debug_dir": "planner_debug_dir",
        },
        "runtime": {
            "job_dir": "runtime_job_dir",
            "root_dir": "runtime_root_dir",
            "save_model_io": "save_model_io",
            "model_io_dump_dir": "model_io_dump_dir",
        },
        "preprocess": {
            "dpi": "dpi",
            "ckpt": "ckpt",
            "backend": "preprocess_backend",
        },
        "render": {
            "canvas_bg": "render_canvas_bg",
            "png_scale": "render_png_scale",
            "max_rows": "render_max_rows",
            "max_cols": "render_max_cols",
        },
        "agent": {
            "model": "agent_model",
            "base_url": "agent_base_url",
            "api_key": "agent_api_key",
            "temperature": "agent_temperature",
            "form_context_mode": "agent_form_context_mode",
            "include_rendered_html": "agent_include_rendered_html",
            "rendered_html_max_chars": "agent_rendered_html_max_chars",
            "skills": "agent_skills",
            "skills_file": "agent_skills_file",
        },
        "reflect": {
            "enable": "reflect_enable",
            "plugin_enable": "reflect_plugin_enable",
            "max_rounds": "reflect_max_rounds",
            "patch_steps": "reflect_patch_steps",
            "checker_model": "reflect_checker_model",
            "edit_log_jsonl": "reflect_edit_log_jsonl",
            "risk_threshold": "reflect_risk_threshold",
            "expand_left_cols": "reflect_expand_left_cols",
            "expand_right_cols": "reflect_expand_right_cols",
            "expand_top_rows": "reflect_expand_top_rows",
            "expand_bottom_rows": "reflect_expand_bottom_rows",
            "focus_pad_left_cols": "reflect_focus_pad_left_cols",
            "focus_pad_top_rows": "reflect_focus_pad_top_rows",
            "focus_pad_right_cols": "reflect_focus_pad_right_cols",
            "focus_pad_bottom_rows": "reflect_focus_pad_bottom_rows",
        },
        "controller": {
            "max_cycles": "controller_max_cycles",
        },
        "source": {
            "instruction": "instruction",
            "manifest": "source_manifest",
            "inline_max_chars": "source_inline_max_chars",
            "pdf_max_pages": "source_pdf_max_pages",
        },
    }
    for key, value in raw.items():
        if key in valid_top_level and not isinstance(value, dict):
            out[key] = value
    for section, mapping in section_mappings.items():
        payload = raw.get(section)
        if not isinstance(payload, dict):
            continue
        for key, dest in mapping.items():
            if key in payload:
                out[dest] = payload[key]
    return out


def _load_config_defaults(config_path: Path, *, explicit: bool) -> dict[str, Any]:
    if not config_path.exists():
        if explicit:
            raise legacy.PipelineError(f"Config file not found: {config_path}")
        return {}
    try:
        import yaml  # type: ignore
    except Exception as exc:
        raise legacy.PipelineError("PyYAML is required to load config.yaml") from exc
    try:
        payload = yaml.safe_load(config_path.read_text(encoding="utf-8")) or {}
    except Exception as exc:
        raise legacy.PipelineError(f"Failed to parse config file {config_path}: {exc}") from exc
    return _flatten_config(payload)


def _collect_unresolved_ranges_from_verification(state: ControllerState) -> list[str]:
    verification = state.verification
    if verification is None:
        return []
    ranges: list[str] = []
    for rng in verification.bad_ranges or []:
        text = str(rng or "").strip()
        if text:
            ranges.append(text)
    for assess in verification.high_risk_cells or []:
        if not isinstance(assess, dict):
            continue
        for key in ("target_range", "target_cell", "candidate_value_cell"):
            text = str(assess.get(key) or "").strip()
            if text:
                ranges.append(text)
    return list(dict.fromkeys(ranges))


def _stalled_retry_regions(state: ControllerState) -> set[str]:
    if state.current_reflect is None or state.verification is None:
        return set()
    unresolved_ranges = _collect_unresolved_ranges_from_verification(state)
    stalled: set[str] = set()
    for item in state.current_reflect.regions:
        if int(state.region_noop_counts.get(item.region, 0) or 0) <= 0:
            continue
        if not unresolved_ranges:
            if state.verification.has_high_risk:
                stalled.add(item.region)
            continue
        for rng in unresolved_ranges:
            try:
                if legacy._range_intersects(item.region, rng):
                    stalled.add(item.region)
                    break
            except Exception:
                continue
    return stalled


def _build_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(description="Planner-driven spreadsheet filling agent.")
    ap.add_argument("--config", default=str(DEFAULT_CONFIG_PATH), help="Path to YAML config file.")
    ap.add_argument("--input", required=True, help="Input file: Excel (.xlsx/.xls), Word (.docx), PDF, or image.")
    ap.add_argument("--instruction", default="Fill the form from the provided source.", help="Natural language instruction")
    ap.add_argument(
        "--source_file",
        action="append",
        default=[],
        help="Additional source document used as filling evidence. Repeat for multiple files.",
    )
    ap.add_argument(
        "--source_manifest",
        default="",
        help="Optional json/txt manifest listing source files used as filling evidence.",
    )
    ap.add_argument(
        "--source_inline_max_chars",
        type=int,
        default=24000,
        help="Max chars of normalized source bundle to inline into the synthesized instruction.",
    )
    ap.add_argument(
        "--source_pdf_max_pages",
        type=int,
        default=30,
        help="Max PDF pages to parse for each source_file PDF. 0 means no page cap.",
    )
    ap.add_argument("--planner_json_path", default="", help="Optional structured planner policy JSON.")
    ap.add_argument(
        "--agentic_planner_enable",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Generate planner policy before running the controller.",
    )
    ap.add_argument("--planner_model", default="", help="Planner model. Defaults to agent_model.")
    ap.add_argument("--planner_base_url", default="", help="Planner base URL. Defaults to agent_base_url.")
    ap.add_argument("--planner_api_key", default="", help="Planner API key. Defaults to agent_api_key.")
    ap.add_argument("--planner_temperature", type=float, default=1.0, help="Planner temperature.")
    ap.add_argument("--planner_debug_dir", default="", help="Optional planner debug directory.")
    ap.add_argument("--save_model_io", action="store_true", help="Save per-call model input/output logs.")
    ap.add_argument("--model_io_dump_dir", default="", help="Optional model I/O dump directory.")
    ap.add_argument("--runtime_job_dir", default="", help="Optional runtime work directory for intermediate artifacts.")
    ap.add_argument(
        "--runtime_root_dir",
        default="",
        help="Alias for --runtime_job_dir when invoking my-pipeline.py directly.",
    )
    ap.add_argument("--output", required=True, help="Output filled Excel")
    ap.add_argument("--sheet", default=None, help="Sheet name (default: active)")
    ap.add_argument("--dpi", type=int, default=96)
    ap.add_argument("--ckpt", default=str(DEFAULT_SEGFORMER_CKPT), help="SegFormer ckpt path")
    ap.add_argument("--preprocess_backend", default="html_css", choices=["html_css", "excel_com"])
    ap.add_argument("--render_canvas_bg", default="#f5f5f5", help="Background color for html_css render.")
    ap.add_argument("--render_png_scale", type=float, default=1.0, help="Scale factor for html_css screenshot.")
    ap.add_argument("--render_max_rows", type=int, default=400, help="Max rows to render from used-range.")
    ap.add_argument("--render_max_cols", type=int, default=120, help="Max cols to render from used-range.")
    ap.add_argument("--agent_model", default="", help="Agent/VLM model name.")
    ap.add_argument("--agent_base_url", default="")
    ap.add_argument("--agent_api_key", default="")
    ap.add_argument("--agent_temperature", type=float, default=1.0)
    ap.add_argument(
        "--agent_form_context_mode",
        default="struct",
        choices=["struct", "html", "struct_html", "refill_html"],
        help="Primary table-understanding context for fill steps.",
    )
    ap.add_argument("--agent_include_rendered_html", action="store_true", help="Include rendered HTML in agent prompts.")
    ap.add_argument("--agent_rendered_html_max_chars", type=int, default=200000, help="Max characters of rendered HTML to append to agent prompts. 0 disables truncation.")
    ap.add_argument("--tool_mode", default="json", choices=["json", "native", "code"])
    ap.add_argument("--max_steps", type=int, default=10)
    ap.add_argument("--external_qwen_pairs_json", default="", help="Optional precomputed slot-pairs JSON path.")
    ap.add_argument("--agent_skills", default="")
    ap.add_argument(
        "--agent_skills_file",
        default="skills.txt",
        help=f"Optional skills text/json file. Relative paths resolve from {DEFAULT_AGENT_SKILLS_DIR}.",
    )
    ap.add_argument(
        "--reflect_enable",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Enable planner-driven reflect and repair.",
    )
    ap.add_argument(
        "--reflect_plugin_enable",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Enable reflect plugin style assessment.",
    )
    ap.add_argument("--reflect_max_rounds", type=int, default=1)
    ap.add_argument("--reflect_patch_steps", type=int, default=6)
    ap.add_argument("--reflect_checker_model", default="")
    ap.add_argument("--reflect_edit_log_jsonl", default="")
    ap.add_argument("--reflect_risk_threshold", type=int, default=1)
    ap.add_argument("--reflect_expand_left_cols", type=int, default=3)
    ap.add_argument("--reflect_expand_right_cols", type=int, default=3)
    ap.add_argument("--reflect_expand_top_rows", type=int, default=1)
    ap.add_argument("--reflect_expand_bottom_rows", type=int, default=1)
    ap.add_argument("--reflect_focus_pad_left_cols", type=int, default=5)
    ap.add_argument("--reflect_focus_pad_top_rows", type=int, default=2)
    ap.add_argument("--reflect_focus_pad_right_cols", type=int, default=3)
    ap.add_argument("--reflect_focus_pad_bottom_rows", type=int, default=2)
    ap.add_argument("--controller_max_cycles", type=int, default=2, help="Maximum planner reflect/repair cycles.")
    return ap


def main() -> None:
    argv = sys.argv[1:]
    config_ns, _ = _build_config_parser().parse_known_args(argv)
    explicit_config = any(arg == "--config" or str(arg).startswith("--config=") for arg in argv)
    config_path = Path(config_ns.config).expanduser()
    if not config_path.is_absolute():
        config_path = (Path.cwd() / config_path).resolve()
    config_defaults = _load_config_defaults(config_path, explicit=explicit_config)
    parser = _build_parser()
    if config_defaults:
        parser.set_defaults(**config_defaults)
    args = parser.parse_args(argv)
    args.config = str(config_path)
    if str(getattr(args, "runtime_root_dir", "") or "").strip() and not str(args.runtime_job_dir or "").strip():
        args.runtime_job_dir = str(args.runtime_root_dir).strip()
    raw_instruction = str(getattr(args, "instruction", "") or "")
    source_files = list(getattr(args, "source_file", []) or [])
    source_manifest = str(getattr(args, "source_manifest", "") or "").strip()
    if not raw_instruction.strip() and not source_files and not source_manifest:
        raise legacy.PipelineError("Provide --instruction or at least one --source_file/--source_manifest.")
    if args.preprocess_backend != "html_css":
        _log(f"[WARN] preprocess_backend={args.preprocess_backend} is deprecated here; forcing html_css-style flow.")
    if args.reflect_enable and not args.reflect_plugin_enable:
        _log("[WARN] planner-driven controller requires reflect_plugin_enable; treating it as enabled.")
        args.reflect_plugin_enable = True
    if config_defaults:
        _log(f"[INFO] loaded config defaults from: {args.config}")
    args.initial_hint_mode = "off"
    args.agent_mapping_mode = "hint"
    args.skip_initial_fill = False

    paths = build_runtime_paths(args, _log)
    args.original_instruction = raw_instruction
    source_bundle = build_source_bundle(
        source_files=source_files,
        source_manifest_path=source_manifest,
        original_instruction=raw_instruction,
        out_dir=paths.job_dir / "source_adapter",
        log_fn=_log,
        max_inline_chars=int(getattr(args, "source_inline_max_chars", 24000) or 0),
        pdf_max_pages=int(getattr(args, "source_pdf_max_pages", 30) or 0),
    )
    args.source_bundle = source_bundle
    if source_files:
        _log(f"[SOURCE_FILES][RAW] count={len(source_files)}")
        for idx, path in enumerate(source_files, start=1):
            _log(f"[SOURCE_FILES][RAW][{idx}] {path}")
    if source_manifest:
        _log(f"[SOURCE_FILES][MANIFEST] {source_manifest}")
    if source_bundle is not None:
        args.instruction = source_bundle.synthesized_instruction
        _log(
            f"[SOURCE_ADAPTER] enabled sources={len(source_bundle.source_files)} "
            f"instruction_chars={len(args.instruction)}"
        )
        for idx, item in enumerate(source_bundle.source_files, start=1):
            _log(f"[SOURCE_FILES][ADAPTED][{idx}] {item}")
    else:
        args.instruction = raw_instruction
    skills_text = load_skills_text(args.agent_skills, args.agent_skills_file)
    form = load_initial_form(args, paths, _log)
    planner_policy, workflow_policy, planner_debug_root = load_planner_policy(args, paths, _log)
    paths.planner_debug_root = planner_debug_root
    state = ControllerState(
        paths=paths,
        skills_text=skills_text,
        form=form,
        planner_policy=planner_policy,
        workflow_policy=workflow_policy,
        usage_summary=create_usage_summary(),
        skip_initial_fill=False,
    )

    bootstrap_plan = build_bootstrap_plan(skip_initial_fill=False)
    _write_json(paths.controller_debug_dir / "bootstrap_plan.json", bootstrap_plan)

    if not args.reflect_enable:
        free_fill(state, args, _log)
        summary = finalize_usage_summary(state)
        _log(f"[USAGE] summary={json.dumps(summary['grand_total'], ensure_ascii=False)} path={paths.job_dir / 'run_usage_summary.json'}")
        _log(f"Done. Filled file: {args.output}")
        return

    max_cycles = max(1, int(args.controller_max_cycles or args.reflect_max_rounds))
    # Keep the controller loop tight by default so it does not spend too many
    # actions on repeated reflect/verify cycles before returning a best-effort file.
    max_actions = max(6, max_cycles * 4 + 2)
    action_count = 0
    while action_count < max_actions and not state.delivered:
        action = plan_next_action(state)
        action_count += 1
        _log(
            f"[AGENT][PLAN] action={action.next_action} region={action.region or '-'} "
            f"use_hints={action.use_hints} reason={action.reason}"
        )
        record: dict[str, Any] = {
            "index": action_count,
            "action": action.next_action,
            "region": action.region,
            "reason": action.reason,
            "metadata": action.metadata,
        }
        if action.next_action == "free_fill":
            result = free_fill(state, args, _log)
            record["result"] = {"dirty": bool(result.get("dirty")), "write_log": len(result.get("write_log", []) or [])}
        elif action.next_action == "reflect_and_hint":
            state.controller_round += 1
            state.regions_restored.clear()
            state.regions_refilled.clear()
            result = reflect_and_hint(state, args, _log, round_idx=state.controller_round)
            record["result"] = asdict(result)
        elif action.next_action == "restore_template_region":
            patch_region = str(action.metadata.get("patch_region") or action.region)
            region_plan = _find_region_plan(state, patch_region)
            preserve_cells = set()
            for rng in region_plan.correct_ranges:
                preserve_cells.update(legacy._iter_range_cells(rng))
            result = restore_template_region(
                template_xlsx=paths.before_fill_xlsx,
                target_xlsx=paths.output_xlsx,
                sheet_name=args.sheet,
                region=action.region,
                preserve_cells=preserve_cells,
            )
            state.regions_restored.add(patch_region)
            record["result"] = result
        elif action.next_action == "guided_refill":
            region_plan = _find_region_plan(state, action.region)
            result = guided_refill(
                state,
                args,
                _log,
                region_plan,
                force_repair=bool(action.metadata.get("force_repair")),
                noop_count=int(action.metadata.get("noop_count", 0) or 0),
            )
            record["result"] = {
                "dirty": bool(result.get("dirty")),
                "write_log": len(result.get("write_log", []) or []),
                "force_repair": bool(result.get("force_repair")),
                "reset_restored_count": int((result.get("reset_result") or {}).get("restored_count", 0) or 0),
                "noop_count_after": int(result.get("noop_count_after", 0) or 0),
            }
        elif action.next_action == "verify_and_deliver":
            result = verify_and_deliver(state, args, _log, round_idx=max(1, state.controller_round))
            record["result"] = asdict(result)
            if not state.delivered:
                stalled_regions = _stalled_retry_regions(state)
                if stalled_regions:
                    _log(
                        "[VERIFY] high-risk regions remain after a no-op repair; retrying stalled regions without a fresh reflect pass: "
                        f"{json.dumps(sorted(stalled_regions), ensure_ascii=False)}"
                    )
                    for region in stalled_regions:
                        state.regions_refilled.discard(region)
                else:
                    _log("[VERIFY] high-risk regions remain; rebuilding local repair plan directly from verification before asking for another full reflect pass.")
                    state.controller_round += 1
                    rebuilt = rebuild_from_verification(state, args, _log, round_idx=state.controller_round)
                    record["verification_rebuild"] = {
                        "round_idx": rebuilt.round_idx,
                        "regions": len(rebuilt.regions),
                        "bad_ranges": len(rebuilt.bad_ranges),
                    }
                    state.regions_restored.clear()
                    state.regions_refilled.clear()
        elif action.next_action == "deliver":
            state.delivered = True
            record["result"] = {"status": "already_delivered"}
        else:
            raise legacy.PipelineError(f"Unknown planner action: {action.next_action}")
        _record_action(state, record)

    if not state.delivered:
        _log(f"[WARN] controller stopped after {action_count} actions without a clean verification pass; delivering best effort output.")

    summary = finalize_usage_summary(state)
    if args.reflect_enable and args.reflect_plugin_enable:
        _log(f"[REFLECT_PLUGIN][FINAL] stats={json.dumps(state.usage_summary['reflect_plugin_stats'], ensure_ascii=False)}")
    _log(f"[USAGE] summary={json.dumps(summary['grand_total'], ensure_ascii=False)} path={paths.job_dir / 'run_usage_summary.json'}")
    _log(f"Done. Filled file: {args.output}")


if __name__ == "__main__":
    main()
