from __future__ import annotations

from typing import Any

try:
    from .pipeline_agent_types import ControllerAction, ControllerState, ReflectRegionPlan
except ImportError:
    from pipeline_agent_types import ControllerAction, ControllerState, ReflectRegionPlan  # type: ignore


def build_bootstrap_plan(skip_initial_fill: bool) -> list[dict[str, Any]]:
    plan: list[dict[str, Any]] = []
    if not skip_initial_fill:
        plan.append(
            {
                "step": "free_fill",
                "goal": "Run an unconstrained first pass with instruction + workbook + shared skills only.",
            }
        )
    plan.append(
        {
            "step": "reflect_and_hint",
            "goal": "Assess the filled workbook, localize bad ranges, and build region hints.",
        }
    )
    plan.append(
        {
            "step": "planner_loop",
            "goal": "For each risky region, choose restore_template_region or guided_refill, then verify again.",
        }
    )
    plan.append(
        {
            "step": "verify_and_deliver",
            "goal": "Stop only when the latest verification shows no remaining high-risk regions.",
        }
    )
    return plan


def _select_next_region(state: ControllerState) -> ReflectRegionPlan | None:
    if state.current_reflect is None:
        return None
    for region_plan in state.current_reflect.regions:
        if region_plan.region not in state.regions_refilled:
            return region_plan
    return None


def plan_next_action(state: ControllerState) -> ControllerAction:
    if state.delivered:
        return ControllerAction(next_action="deliver", reason="Pipeline is already marked delivered.")

    if not state.skip_initial_fill and not state.first_pass_result:
        return ControllerAction(
            next_action="free_fill",
            reason="No first-pass workbook exists yet, so the controller must start with an unconstrained fill.",
        )

    if state.current_reflect is None:
        if state.verification is not None and state.verification.has_high_risk:
            return ControllerAction(
                next_action="reflect_and_hint",
                reason="Verification still reports high-risk regions, so the controller needs a new hinted assessment pass.",
            )
        return ControllerAction(
            next_action="reflect_and_hint",
            reason="The controller needs a structured assessment before it can choose restore or refill actions.",
        )

    if (not state.current_reflect.has_high_risk) and (not state.current_reflect.bad_ranges):
        return ControllerAction(
            next_action="verify_and_deliver",
            reason="The latest hinted assessment found no unresolved risky ranges, so the next step is final verification.",
        )

    region_plan = _select_next_region(state)
    if region_plan is not None:
        noop_count = int(state.region_noop_counts.get(region_plan.region, 0) or 0)
        has_direct_hints = bool(region_plan.selected_region_hints)
        has_negative_hints = bool(region_plan.invalid_slots)
        qwen_context = region_plan.qwen_context_summary if isinstance(region_plan.qwen_context_summary, dict) else {}
        has_qwen_context = bool(
            qwen_context.get("summary_lines")
            or qwen_context.get("candidate_key_groups")
            or qwen_context.get("negative_only_key_groups")
        )
        force_repair = (not has_direct_hints) or (noop_count > 0)
        if region_plan.need_restore_template and region_plan.region not in state.regions_restored:
            restore_region = region_plan.restore_template_region or region_plan.region
            reason = "; ".join(region_plan.restore_reasons) or "template/static cells appear corrupted"
            return ControllerAction(
                next_action="restore_template_region",
                region=restore_region,
                reason=reason,
                metadata={"patch_region": region_plan.region},
            )
        return ControllerAction(
            next_action="guided_refill",
            region=region_plan.region,
            use_hints=(has_direct_hints or has_negative_hints or has_qwen_context),
            reason=(
                "local region hints are available and the region is ready for a guarded repair pass"
                if has_direct_hints and noop_count <= 0
                else (
                    "no direct positive hints survived selection, but Qwen local context is available, so the controller is escalating to a forced repair pass with soft regional evidence"
                    if has_qwen_context and noop_count <= 0
                    else (
                        "only negative / invalid-slot hints are available, so the controller is escalating it to a forced cell-level review repair"
                        if noop_count <= 0
                        else "the previous guided repair made no edits, so the controller is escalating this region to a forced repair retry"
                    )
                )
            ),
            metadata={
                "force_repair": force_repair,
                "noop_count": noop_count,
                "has_direct_hints": has_direct_hints,
                "has_negative_hints": has_negative_hints,
                "has_qwen_context": has_qwen_context,
            },
        )

    return ControllerAction(
        next_action="verify_and_deliver",
        reason="All currently known risky regions were already handled, so the controller should verify before looping again.",
    )
