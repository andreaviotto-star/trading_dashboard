#!/usr/bin/env python3
"""Consolidated audit runner for the Quant Portfolio Dashboard.

The audit is intentionally methodology-preserving. It checks source-level
invariants that can be run without Streamlit/data dependencies, and optionally
runs the existing Golden Master when its supporting files are present.
"""
from __future__ import annotations

import argparse
import ast
import json
import re
import subprocess
import sys
from pathlib import Path


def run(cmd: list[str], cwd: Path) -> tuple[int, str]:
    p = subprocess.run(cmd, cwd=cwd, text=True, capture_output=True)
    return p.returncode, ((p.stdout or "") + (p.stderr or "")).strip()


def parse_app(app_path: Path) -> tuple[str, ast.Module, str]:
    text = app_path.read_text(encoding="utf-8")
    tree = ast.parse(text)
    # The header is stale in the historical v25.1 file. Prefer the explicit
    # v25 feature markers when no APP_VERSION constant exists.
    m = re.search(r"APP_VERSION\s*=\s*['\"]([^'\"]+)", text)
    if m:
        version = m.group(1)
    elif all(x in text for x in ("portfolio_risk", "Ledoit-Wolf", "CARP")):
        version = "v25.x (header marker stale)"
    else:
        version = "unknown"
    return text, tree, version


def check_required_symbols(tree: ast.Module) -> tuple[bool, str]:
    funcs = {n.name for n in tree.body if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef))}
    required = {
        "load_all_systems", "compute_metrics", "compute_portfolio_metrics",
        "portfolio_risk", "compute_risk_parity_sizing",
        "compute_correlation_aware_rp", "monte_carlo_simulation",
    }
    missing = sorted(required - funcs)
    return not missing, "all required quantitative functions found" if not missing else "missing: " + ", ".join(missing)


def check_startup_sizing(text: str) -> tuple[bool, str]:
    fn = text.find("def _render_tab3():")
    if fn < 0:
        return False, "_render_tab3 not found"
    block = text[fn:fn + 18000]
    init = block.find("_current_sizing_map =")
    first_use = min([x for x in (block.find("_build_periodic_trade_pnl", init + 1), block.find("portfolio_risk", init + 1)) if x >= 0] or [10**9])
    if init < 0 or first_use == 10**9:
        return False, "could not locate sizing-map initialization/use"
    return init < first_use, "_current_sizing_map is initialized before Tab 3 risk/correlation use"


def check_psd_tolerance(text: str) -> tuple[bool, str]:
    expected = "if _psd_correction > 1e-6:"
    return expected in text, "PSD warning ignores corrections <= 1e-6" if expected in text else "material PSD tolerance not found"


def check_carp_methodology(text: str) -> tuple[bool, str]:
    required = [
        "integer = np.clip(np.rint(continuous).astype(int), 0, max_contracts)",
        "active = q > 0; desired = 1.0 / active.sum()",
        "np.clip(cand[i] + delta, 0, max_contracts)",
    ]
    missing = [x for x in required if x not in text]
    if missing:
        return False, "CARP methodology changed unexpectedly: " + "; ".join(missing)
    return True, "established CARP system-selection methodology preserved (zero contracts remain allowed)"


def check_quantitative_markers(text: str) -> tuple[bool, str]:
    markers = [
        "def portfolio_risk(",
        "def build_correlation_daily_pnl(",
        "def _pairwise_corr_cov(",
        "Ledoit-Wolf",
        "risk_contribution",
    ]
    missing = [x for x in markers if x not in text]
    return not missing, "canonical risk/correlation layer present" if not missing else "missing: " + "; ".join(missing)


def compare_golden(root: Path, require: bool = False) -> tuple[str, str]:
    golden = root / "golden_master.py"
    baseline = root / "baseline.csv"
    diff_script = root / "diff_golden_master.py"
    if not (golden.exists() and baseline.exists() and diff_script.exists()):
        detail = "Golden Master snapshot skipped (baseline/harness not committed to this branch)"
        return ("FAIL" if require else "SKIP", detail)
    post = root / ".audit_post.csv"
    rc, out = run([sys.executable, str(golden), "--source", "app.py", "--output", str(post)], root)
    if rc:
        return "FAIL", "Golden Master failed: " + out[-1200:]
    rc, out = run([sys.executable, str(diff_script), str(baseline), str(post)], root)
    if rc:
        return "FAIL", "Golden Master comparison failed: " + out[-1200:]
    return "PASS", "Golden Master generated and comparison completed"


def write_report(root: Path, checks: list[dict], version: str) -> Path:
    failed = sum(c["status"] == "FAIL" for c in checks)
    status = "FAIL" if failed else "PASS"
    lines = [f"# Quant Portfolio Dashboard Audit — {version}", "", f"**Overall:** {status}", "", "| Check | Status | Detail |", "|---|---|---|"]
    for c in checks:
        lines.append(f"| {c['name']} | {c['status']} | {c['detail'].replace(chr(10), ' ')} |")
    path = root / "audit_report.md"
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", default=".")
    ap.add_argument("--require-golden", action="store_true", help="fail if the Golden Master baseline/harness is unavailable")
    args = ap.parse_args()
    root = Path(args.root).resolve()
    app = root / "app.py"
    if not app.exists():
        print("FAIL: app.py not found")
        return 1

    checks: list[dict] = []
    try:
        text, tree, version = parse_app(app)
        checks.append({"name": "Python syntax", "status": "PASS", "detail": "app.py parses successfully"})
    except SyntaxError as exc:
        print(f"FAIL: app.py syntax error: {exc}")
        return 1

    for name, fn in [
        ("Required quantitative functions", lambda: check_required_symbols(tree)),
        ("Canonical risk/correlation layer", lambda: check_quantitative_markers(text)),
        ("Tab 3 sizing-map startup order", lambda: check_startup_sizing(text)),
        ("PSD numerical tolerance", lambda: check_psd_tolerance(text)),
        ("CARP methodology preserved", lambda: check_carp_methodology(text)),
    ]:
        try:
            ok, detail = fn()
            checks.append({"name": name, "status": "PASS" if ok else "FAIL", "detail": detail})
        except Exception as exc:
            checks.append({"name": name, "status": "FAIL", "detail": str(exc)})

    golden_status, golden_detail = compare_golden(root, args.require_golden)
    checks.append({"name": "Golden Master", "status": golden_status, "detail": golden_detail})

    report = write_report(root, checks, version)
    payload = {"version": version, "status": "FAIL" if any(c["status"] == "FAIL" for c in checks) else "PASS", "checks": checks}
    (root / "audit_result.json").write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")

    print(f"\nQuant Portfolio Dashboard Audit — {version}\n" + "=" * 52)
    for c in checks:
        print(f"{c['status']:4}  {c['name']}: {c['detail']}")
    print("=" * 52 + f"\nOverall: {payload['status']}\nReport: {report}")
    return 0 if payload["status"] == "PASS" else 1


if __name__ == "__main__":
    raise SystemExit(main())
