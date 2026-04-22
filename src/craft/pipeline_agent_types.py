from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass
class PipelinePaths:
    job_dir: Path
    source_input: Path
    input_xlsx: Path
    before_fill_xlsx: Path
    output_xlsx: Path
    edit_log_jsonl: Path
    controller_debug_dir: Path
    input_plan_json: Path | None = None
    model_io_dir: Path | None = None
    planner_debug_root: Path | None = None


@dataclass
class ReflectRegionPlan:
    region: str
    patch_targets: list[dict[str, Any]] = field(default_factory=list)
    selected_region_hints: list[dict[str, Any]] = field(default_factory=list)
    invalid_slots: list[dict[str, Any]] = field(default_factory=list)
    qwen_context_summary: dict[str, Any] = field(default_factory=dict)
    correct_ranges: list[str] = field(default_factory=list)
    need_restore_template: bool = False
    restore_template_region: str = ""
    restore_reasons: list[str] = field(default_factory=list)
    raw_region_hints: list[dict[str, Any]] = field(default_factory=list)
    raw_patch_job: dict[str, Any] = field(default_factory=dict)


@dataclass
class ReflectStageResult:
    round_idx: int = 0
    stage_name: str = "reflect_and_hint"
    screenshot_check: dict[str, Any] = field(default_factory=dict)
    reflect_result: dict[str, Any] = field(default_factory=dict)
    bad_ranges: list[str] = field(default_factory=list)
    correct_ranges: list[str] = field(default_factory=list)
    regions: list[ReflectRegionPlan] = field(default_factory=list)
    has_high_risk: bool = False
    high_risk_cells: list[dict[str, Any]] = field(default_factory=list)
    missing_keys: list[str] = field(default_factory=list)
    assess_summary: dict[str, Any] = field(default_factory=dict)


@dataclass
class ControllerAction:
    next_action: str
    reason: str
    region: str = ""
    use_hints: bool = False
    metadata: dict[str, Any] = field(default_factory=dict)


@dataclass
class ControllerState:
    paths: PipelinePaths
    skills_text: str
    form: dict[str, Any]
    planner_policy: dict[str, Any]
    workflow_policy: dict[str, Any]
    usage_summary: dict[str, Any]
    skip_initial_fill: bool = False
    first_pass_result: dict[str, Any] = field(default_factory=dict)
    first_pass_sheet_png: str = ""
    current_reflect: ReflectStageResult | None = None
    verification: ReflectStageResult | None = None
    action_history: list[dict[str, Any]] = field(default_factory=list)
    regions_restored: set[str] = field(default_factory=set)
    regions_refilled: set[str] = field(default_factory=set)
    region_noop_counts: dict[str, int] = field(default_factory=dict)
    controller_round: int = 0
    verification_attempts: int = 0
    delivered: bool = False
