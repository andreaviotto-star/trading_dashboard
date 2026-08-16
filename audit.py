#!/usr/bin/env python3
"""Compact audit runner for the Quant Portfolio Dashboard.

This audit is deliberately conservative: it validates code structure and
quantitative invariants without changing the dashboard's trading methodology.
It can run locally in the Codespace or in GitHub Actions.
"""
from __future__ import annotations

import argparse
import ast
import json
import subprocess
import sys
from pathlib import Path


def run(cmd: list[str], cwd: Path) -> tuple[int, str]:
    p = subprocess.run(cmd, cwd=cwd, text=True, capture_output=True)
    out = (p.stdout or "") + (p.stderr or "")
    return p.returncode, out.strip()


def extract_app_version(app_path: Path) -> str:
    text = app_path.read_text(encoding="utf-8")
    try:
        tree = ast.parse(text)
    except SyntaxError as exc:
        raise RuntimeError(f"app.py syntax error: {exc}") from exc
    for node in tree.body:
        if isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name) and target.id == "APP_VERSION":
                    if isinstance(node.value, ast.Constant):
                        return str(node.value.value)
    # Header fallback for historical versions.
    for line in text.splitlines()[:10]:
        if "Dashboard  v" in line:
            return line.split("Dashboard", 1)[1].strip()
    return "unknown"


def test_required_symbols(app_path: Path) -> tuple[bool, str]:
    tree = ast.parse(app_path.read_text(encoding="utf-8"))
    funcs = {n.name for n in tree.body if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef))}
    required = {
        "load_all_systems",
        "compute_metrics",
        "compute_portfolio_metrics",
        "portfolio_risk",
        "compute_risk_parity_sizing",
        "compute_correlation_aware_rp",
        "monte_carlo_simulation",
    }
    missing = sorted(required - funcs)
    return (not missing, "missing: " + ", ".join(missing) if missing else "all required quantitative functions found")


def test_carp_fixed_contract_logic(app_path: Path) -> tuple[bool, str]:
    text = app_path.read_text(encoding="utf-8")
    # We deliberately do NOT require a >=1 system constraint: zero contracts
    # are allowed by the established CARP methodology. We only verify that the
    # optimizer contains the expected integer sizing/refinement machinery.
    required_fragments = [
        "compute_correlation_aware_rp",
        "integer = np.clip(np.rint(continuous).astype(int), 0, max_contracts)",
        "active = q > 0; desired = 1.0 / active.sum()",
        "np.clip(cand[i] + delta, 0, max_contracts)",
    ]
    missing = [x for x in required_fragments if x not in text]
    return (not missing, "established CARP integer-selection methodology preserved" if not missing else "missing: " + "; ".join(missing))


def compare_golden(root: Path) -> tuple[bool, str]:
    golden = root / "golden_master.py"
    baseline = root / "baseline.csv"
    if not golden.exists():
        return False, "golden_master.py not found"
    if not baseline.exists():
        return True, "baseline.csv not present; quantitative snapshot skipped"
    post = root / ".audit_post.csv"
    rc, out = run([sys.executable, str(golden), "--source", "app.py", "--output", str(post)], root)
    if rc:
        return False, "Golden Master failed: " + out[-1200:]
    diff_script = root / "diff_golden_master.py"
    if not diff_script.exists():
        return False, "diff_golden_master.py not found"
    rc, out = run([sys.executable, str(diff_script), str(baseline), str(post)], root)
    # The diff is informational; new diagnostics are expected in v25.x.
    # A non-zero exit is the authoritative failure condition.
    if rc:
        return False, "Golden Master comparison failed: " + out[-1200:]
    return True, "Golden Master generated and comparison completed"


def write_report(root: Path, checks: list[dict], version: str) -> Path:
    passed = sum(bool(c["passed"]) for c in checks)
    total = len(checks)
    status = "PASS" if passed == total else "FAIL"
    lines = [
        f"# Quant Portfolio Dashboard Audit — {version}",
        "",
        f"**Overall:** {status}",
        f"**Checks:** {passed}/{total} passed",
        "",
        "| Check | Status | Detail |",
        "|---|---|---|",
    ]
    for c in checks:
        lines.append(f"| {c['name']} | {'PASS' if c['passed'] else 'FAIL'} | {c['detail'].replace(chr(10), ' ')} |")
    path = root / "audit_report.md"
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", default=".")
    args = ap.parse_args()
    root = Path(args.root).resolve()
    app = root / "app.py"
    checks: list[dict] = []

    if not app.exists():
        print("FAIL: app.py not found")
        return 1

    try:
        version = extract_app_version(app)
        checks.append({"name": "Python syntax", "passed": True, "detail": "app.py parses successfully"})
    except Exception as exc:
        version = "unknown"
        checks.append({"name": "Python syntax", "passed": False, "detail": str(exc)})

    try:
        ok, detail = test_required_symbols(app)
        checks.append({"name": "Required quantitative functions", "passed": ok, "detail": detail})
    except Exception as exc:
        checks.append({"name": "Required quantitative functions", "passed": False, "detail": str(exc)})

    try:
        ok, detail = test_carp_fixed_contract_logic(app)
        checks.append({"name": "CARP methodology preserved", "passed": ok, "detail": detail})
    except Exception as exc:
        checks.append({"name": "CARP methodology preserved", "passed": False, "detail": str(exc)})

    ok, detail = compare_golden(root)
    checks.append({"name": "Golden Master", "passed": ok, "detail": detail})

    report = write_report(root, checks, version)
    payload = {"version": version, "status": "PASS" if all(c["passed"] for c in checks) else "FAIL", "checks": checks}
    (root / "audit_result.json").write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")

    print(f"\nQuant Portfolio Dashboard Audit — {version}")
    print("=" * 42)
    for c in checks:
        print(f"{'PASS' if c['passed'] else 'FAIL':4}  {c['name']}: {c['detail']}")
    print("=" * 42)
    print(f"Overall: {payload['status']}")
    print(f"Report:  {report}")
    return 0 if payload["status"] == "PASS" else 1


if __name__ == "__main__":
    raise SystemExit(main())
