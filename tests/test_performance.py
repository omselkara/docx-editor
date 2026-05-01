"""Performance benchmarks using pytest-benchmark.

Usage:
  pytest tests/test_performance.py -v
  pytest tests/test_performance.py --benchmark-save=baseline
  pytest tests/test_performance.py --benchmark-compare=baseline
  pytest tests/test_performance.py --benchmark-histogram

Install: pip install pytest-benchmark
"""
from __future__ import annotations

import os
import sys
import json

import pytest

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../scripts")))

from docx_engine import reading, editing, tables, batch_tools
from docx_engine.core import create_document, load_document

pytest.importorskip("pytest_benchmark", reason="pytest-benchmark not installed; skipping perf tests")


# ─── full_text ────────────────────────────────────────────────────────────────

def bench_full_text_lxml(benchmark, large_doc):
    """LXML fast-path — target < 200ms on 1000-paragraph doc."""
    result = benchmark(reading.full_text, large_doc, include_formatting=False, json_mode=True)
    assert result["status"] == "SUCCESS"


def bench_full_text_with_formatting(benchmark, large_doc):
    """Python-docx slow path (baseline for comparison)."""
    result = benchmark(reading.full_text, large_doc, include_formatting=True, json_mode=True)
    assert result["status"] == "SUCCESS"


# ─── get_outline ─────────────────────────────────────────────────────────────

def bench_outline_large_doc(benchmark, large_doc):
    """Outline extraction on doc with 20 headings — target < 50ms."""
    result = benchmark(reading.get_outline, large_doc, json_mode=True)
    assert result["status"] == "SUCCESS"


# ─── search_text ─────────────────────────────────────────────────────────────

def bench_search_regex(benchmark, large_doc):
    """Regex search across 1000-paragraph doc — target < 300ms."""
    result = benchmark(reading.search_text, large_doc, r"Section \d+", json_mode=True)
    assert result["status"] == "SUCCESS"


# ─── replace_text ────────────────────────────────────────────────────────────

def bench_replace_text_single_hit(benchmark, large_doc, tmp_path):
    """Single-match replace — target < 500ms."""
    import shutil
    copy = str(tmp_path / "replace_bench.docx")
    shutil.copy2(large_doc, copy)

    def _run():
        shutil.copy2(large_doc, copy)
        editing.replace_text(copy, "Section 1", "SECTION_ONE")

    benchmark(_run)


# ─── create_table ────────────────────────────────────────────────────────────

def bench_create_table_10x10(benchmark, tmp_path):
    """Create 10-column, 10-row table — target < 300ms."""
    path = str(tmp_path / "table_bench.docx")
    create_document(path)
    headers = ",".join(f"Col{i}" for i in range(10))
    rows_data = ";".join(",".join(f"R{r}C{c}" for c in range(10)) for r in range(9))

    def _run():
        create_document(path)
        tables.create_table(path, headers=headers, rows_data=rows_data)

    benchmark(_run)


# ─── batch_execute ────────────────────────────────────────────────────────────

def bench_batch_execute_20_ops(benchmark, simple_doc, tmp_path):
    """20-operation batch — target < 500ms."""
    import shutil

    cmds_path = str(tmp_path / "bench_cmds.json")
    commands = [
        {"action": "replace_text", "args": {"find_pattern": "paragraph", "replace_with": "section"}}
    ] * 5 + [
        {"action": "insert_paragraph", "args": {"text": f"Bench paragraph {i}"}}
        for i in range(15)
    ]
    with open(cmds_path, "w") as f:
        json.dump({"on_error": "continue", "commands": commands}, f)

    copy = str(tmp_path / "batch_bench.docx")

    def _run():
        shutil.copy2(simple_doc, copy)
        batch_tools.batch_execute(copy, cmds_path, json_mode=True)

    benchmark(_run)


# ─── LCS diff ────────────────────────────────────────────────────────────────

def bench_diff_documents(benchmark, large_doc, tmp_path):
    """LCS diff on 1000-paragraph docs (capped at 500) — target < 1000ms."""
    import shutil
    copy = str(tmp_path / "diff_bench.docx")
    shutil.copy2(large_doc, copy)

    result = benchmark(batch_tools.diff_documents, large_doc, copy, json_mode=True)
    assert result["status"] == "SUCCESS"
