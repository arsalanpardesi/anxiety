"""
Validate an Anxiety Register JSON file.

Turns the manual Step-3 merge into a checked contract: a malformed finding can no longer
silently corrupt the register or its markdown summary. Stdlib-only, no third-party deps.

Checks:
  - required top-level keys and the three phase buckets
  - each finding has the required keys with valid enum values
  - a finding's `phase` matches the bucket it sits in
  - ids are non-empty and unique across the whole register
  - `auto_fixable` is a boolean
  - phase scores, summary counts, verdict, and overall score are internally consistent
    (severity-weighted: blocking Critical/High deduct in full, advisory Medium/Low capped)

Usage:
  python validate_register.py <name>.anxiety.json

Developed by Arsalan Pardesi. MIT License.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any, Dict, List

sys.path.insert(0, str(Path(__file__).resolve().parent))
try:
    from anxiety_scan import (score_findings, verdict_for, VERDICT_LABELS,
                              ADVISORY_CAP_PHASE, ADVISORY_CAP_OVERALL)
    _HAVE_SCORER = True
except Exception:  # noqa: BLE001 - scoring checks are skipped if the scanner can't be imported
    _HAVE_SCORER = False

PHASES = ("coverage", "correctness", "adversarial")
SEVERITIES = ("critical", "high", "medium", "low")
STATUSES = ("open", "resolved", "dismissed")
AGENT_CHECK_STATUSES = ("pending", "completed", "not_applicable")
FINDING_KEYS = ("id", "phase", "category", "severity", "title", "detail",
                "evidence", "remediation", "auto_fixable", "status")


def validate_register(reg: Any, require_complete: bool = False) -> List[str]:
    """Return a list of human-readable error strings; empty means the register is valid.

    With require_complete=True the Step-2 agent checks must all be completed or marked
    not_applicable (no 'pending'), so a clean-looking register can't be presented before the
    reasoning checks were actually done.
    """
    errors: List[str] = []
    if not isinstance(reg, dict):
        return ["register is not a JSON object"]

    for key in ("anxiety_register_id", "deliverable_path", "phases", "summary"):
        if key not in reg:
            errors.append(f"missing top-level key: '{key}'")

    phases = reg.get("phases")
    if not isinstance(phases, dict):
        errors.append("'phases' must be an object")
        return errors

    seen_ids: Dict[str, int] = {}
    all_findings: List[Dict[str, Any]] = []
    for ph in PHASES:
        if ph not in phases:
            errors.append(f"missing phase: '{ph}'")
            continue
        bucket = phases[ph]
        if not isinstance(bucket, dict):
            errors.append(f"phase '{ph}' must be an object")
            continue
        findings = bucket.get("findings")
        if not isinstance(findings, list):
            errors.append(f"phase '{ph}' has no findings list")
            findings = []
        for i, f in enumerate(findings):
            loc = f"{ph}.findings[{i}]"
            if not isinstance(f, dict):
                errors.append(f"{loc} is not an object")
                continue
            for k in FINDING_KEYS:
                if k not in f:
                    errors.append(f"{loc} missing key '{k}'")
            if f.get("severity") not in SEVERITIES:
                errors.append(f"{loc} invalid severity: {f.get('severity')!r}")
            if f.get("status") not in STATUSES:
                errors.append(f"{loc} invalid status: {f.get('status')!r}")
            if f.get("phase") != ph:
                errors.append(f"{loc} phase '{f.get('phase')}' does not match bucket '{ph}'")
            if not isinstance(f.get("auto_fixable"), bool):
                errors.append(f"{loc} 'auto_fixable' must be a boolean")
            fid = f.get("id")
            if not fid or not isinstance(fid, str):
                errors.append(f"{loc} has an empty or non-string id")
            else:
                seen_ids[fid] = seen_ids.get(fid, 0) + 1
            all_findings.append(f)

        if _HAVE_SCORER:
            score = bucket.get("score")
            expected = score_findings(findings, ADVISORY_CAP_PHASE)
            if score != expected:
                errors.append(f"phase '{ph}' score {score} != recomputed {expected}")

    for fid, count in seen_ids.items():
        if count > 1:
            errors.append(f"duplicate finding id: '{fid}' ({count} times)")

    summary = reg.get("summary")
    if not isinstance(summary, dict):
        errors.append("'summary' must be an object")
        return errors

    counts = {sev: sum(1 for f in all_findings if f.get("severity") == sev) for sev in SEVERITIES}
    active = [f for f in all_findings if f.get("status", "open") != "dismissed"]
    checks = {
        "total_findings": len(all_findings),
        "critical": counts["critical"],
        "high": counts["high"],
        "medium": counts["medium"],
        "low": counts["low"],
        "blocking": sum(1 for f in active if f.get("severity") in ("critical", "high")),
        "advisory": sum(1 for f in active if f.get("severity") in ("medium", "low")),
        "auto_fixable": sum(1 for f in all_findings if f.get("auto_fixable") is True),
    }
    for key, expected in checks.items():
        if key in summary and summary.get(key) != expected:
            errors.append(f"summary.{key} {summary.get(key)} != recomputed {expected}")

    if _HAVE_SCORER:
        if summary.get("overall_confidence") != score_findings(all_findings, ADVISORY_CAP_OVERALL):
            errors.append(
                f"summary.overall_confidence {summary.get('overall_confidence')} != recomputed "
                f"{score_findings(all_findings, ADVISORY_CAP_OVERALL)}")
        expected_verdict = verdict_for(all_findings)
        if "verdict" in summary and summary.get("verdict") != expected_verdict:
            errors.append(f"summary.verdict {summary.get('verdict')!r} != recomputed "
                          f"{expected_verdict!r}")
        if "verdict_label" in summary and summary.get("verdict_label") != VERDICT_LABELS[expected_verdict]:
            errors.append("summary.verdict_label does not match the recomputed verdict")
        for ph in PHASES:
            key = f"{ph}_score"
            if ph in phases and isinstance(phases[ph], dict):
                expected = phases[ph].get("score")
                if summary.get(key) != expected:
                    errors.append(f"summary.{key} {summary.get(key)} != phase score {expected}")

    agent_checks = reg.get("agent_checks")
    if agent_checks is not None:
        if not isinstance(agent_checks, dict):
            errors.append("'agent_checks' must be an object")
        else:
            for name, entry in agent_checks.items():
                if not isinstance(entry, dict) or entry.get("status") not in AGENT_CHECK_STATUSES:
                    errors.append(f"agent_checks.{name} has an invalid or missing status")
    if require_complete:
        if not isinstance(agent_checks, dict) or not agent_checks:
            errors.append("require-complete: 'agent_checks' is missing; the Step-2 agent checks "
                          "were not recorded")
        else:
            pending = [k for k, v in agent_checks.items()
                       if not isinstance(v, dict) or v.get("status") == "pending"]
            for k in pending:
                errors.append(f"require-complete: agent check '{k}' is still pending "
                              "(complete it or mark it not_applicable)")

    return errors


def main(argv: List[str]) -> int:
    import argparse
    ap = argparse.ArgumentParser(description="Validate an anxiety register JSON file.")
    ap.add_argument("register", help="Path to <name>.anxiety.json")
    ap.add_argument("--require-complete", action="store_true",
                    help="Also require every Step-2 agent check to be completed or "
                         "not_applicable (no 'pending'). Use before presenting results.")
    args = ap.parse_args(argv[1:])

    path = Path(args.register)
    if not path.is_file():
        print(f"Error: file not found: {path}")
        return 2
    try:
        reg = json.loads(path.read_text(encoding="utf-8"))
    except (ValueError, OSError) as exc:
        print(f"Error: could not read JSON: {exc}")
        return 2

    errors = validate_register(reg, require_complete=args.require_complete)
    if not errors:
        print(f"OK: {path} is a valid anxiety register.")
        return 0
    print(f"INVALID: {path} has {len(errors)} problem(s):")
    for e in errors:
        print(f"  - {e}")
    return 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
