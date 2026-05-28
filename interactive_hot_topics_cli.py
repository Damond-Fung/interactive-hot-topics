#!/usr/bin/env python3
"""Unified CLI for the interactive hot topics report.

Three sub-commands are provided so that other AI tools without an
agent-callback loop can still complete the second-step AI judgement
via a single shell command:

    report   collect forum data and export the base workbook only
    analyze  run per-topic AI judgement against an already-generated report
    run      one-shot pipeline: collect + AI judgement + final workbook

All sub-commands are thin wrappers that forward to the existing
`interactive_hot_topics_report.py` main entry, so behaviour stays in a
single place and we avoid duplicating logic.
"""

from __future__ import annotations

import argparse
import runpy
import sys
from pathlib import Path


REPO_SCRIPT_NAME = "interactive_hot_topics_report.py"


def locate_repo_script() -> Path:
    here = Path(__file__).resolve()
    for parent in [here.parent, *here.parents]:
        candidate = parent / REPO_SCRIPT_NAME
        if candidate.exists():
            return candidate
    raise FileNotFoundError(f"Unable to locate {REPO_SCRIPT_NAME} next to the CLI.")


def _invoke_report_script(forwarded_argv: list[str]) -> None:
    script_path = locate_repo_script()
    repo_root = script_path.parent
    if str(repo_root) not in sys.path:
        sys.path.insert(0, str(repo_root))
    sys.argv = [str(script_path), *forwarded_argv]
    runpy.run_path(str(script_path), run_name="__main__")


def _add_common_window_args(p: argparse.ArgumentParser) -> None:
    p.add_argument("--forum-base-url")
    p.add_argument("--timezone")
    p.add_argument("--time-preset", choices=["last-week", "this-week", "last-7-days"])
    p.add_argument("--start-date")
    p.add_argument("--end-date")
    p.add_argument("--top-n", type=int)
    p.add_argument("--output")
    p.add_argument("--max-pages", type=int)


def _add_common_llm_args(p: argparse.ArgumentParser) -> None:
    p.add_argument("--llm-provider", choices=["github", "openai", "custom"])
    p.add_argument("--llm-base-url")
    p.add_argument("--llm-model")
    p.add_argument("--llm-api-key-env")
    p.add_argument("--llm-timeout", type=int)
    p.add_argument("--llm-max-retries", type=int)
    p.add_argument("--llm-max-topics", type=int)


def _collect_forwarded(ns: argparse.Namespace, names: list[str]) -> list[str]:
    out: list[str] = []
    for name in names:
        attr = name.lstrip("-").replace("-", "_")
        value = getattr(ns, attr, None)
        if value is None:
            continue
        if isinstance(value, bool):
            if value:
                out.append(name)
            continue
        out.extend([name, str(value)])
    return out


WINDOW_FLAGS = [
    "--forum-base-url",
    "--timezone",
    "--time-preset",
    "--start-date",
    "--end-date",
    "--top-n",
    "--output",
    "--max-pages",
]

LLM_FLAGS = [
    "--llm-provider",
    "--llm-base-url",
    "--llm-model",
    "--llm-api-key-env",
    "--llm-timeout",
    "--llm-max-retries",
    "--llm-max-topics",
]


def cmd_report(ns: argparse.Namespace) -> None:
    """Collect forum data only (Agent-pending), no AI judgement."""
    forwarded = _collect_forwarded(ns, WINDOW_FLAGS)
    forwarded += ["--ai-mode", "agent", "--export-ai-template"]
    _invoke_report_script(forwarded)


def cmd_analyze(ns: argparse.Namespace) -> None:
    """Re-aggregate using an existing AI results JSON file (Agent mode)."""
    if not ns.ai_results:
        raise SystemExit("analyze requires --ai-results <path-to-ai_results.json>")
    forwarded = _collect_forwarded(ns, WINDOW_FLAGS)
    forwarded += ["--ai-results", ns.ai_results, "--ai-mode", "agent"]
    _invoke_report_script(forwarded)


def cmd_run(ns: argparse.Namespace) -> None:
    """One-shot: collect + per-topic AI judgement + final workbook."""
    forwarded = _collect_forwarded(ns, WINDOW_FLAGS)
    forwarded += _collect_forwarded(ns, LLM_FLAGS)
    forwarded += ["--ai-mode", ns.ai_mode or "api", "--export-ai-template"]
    _invoke_report_script(forwarded)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="interactive-hot-topics",
        description=(
            "Interactive hot topics report CLI. Use `run` for a single-shot pipeline "
            "in environments that cannot drive an Agent callback loop."
        ),
    )
    sub = parser.add_subparsers(dest="command", required=True)

    p_report = sub.add_parser("report", help="Collect forum data and export the base workbook only.")
    _add_common_window_args(p_report)
    p_report.set_defaults(func=cmd_report)

    p_analyze = sub.add_parser("analyze", help="Re-aggregate using an Agent-supplied ai_results.json.")
    _add_common_window_args(p_analyze)
    p_analyze.add_argument("--ai-results", required=True)
    p_analyze.set_defaults(func=cmd_analyze)

    p_run = sub.add_parser(
        "run",
        help="One-shot pipeline: collect + AI judgement (API mode) + final workbook.",
    )
    _add_common_window_args(p_run)
    _add_common_llm_args(p_run)
    p_run.add_argument("--ai-mode", choices=["auto", "api"], default="api")
    p_run.set_defaults(func=cmd_run)

    return parser


def main(argv: list[str] | None = None) -> None:
    parser = build_parser()
    ns = parser.parse_args(argv)
    ns.func(ns)


if __name__ == "__main__":
    main()
