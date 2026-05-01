"""Tests for docx_engine.reading — outline, full_text, search, read_range."""
from __future__ import annotations

import os
import sys

import pytest

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../scripts")))

from docx_engine import reading
from docx_engine.core import load_document


class TestGetOutline:
    def test_outline_with_headings(self, heading_doc):
        res = reading.get_outline(heading_doc, json_mode=True)
        assert res["status"] == "SUCCESS"
        assert len(res["headings"]) == 3
        texts = [h["text"] for h in res["headings"]]
        assert "Introduction" in texts
        assert "Methods" in texts
        assert "Results" in texts

    def test_outline_levels(self, heading_doc):
        res = reading.get_outline(heading_doc, json_mode=True)
        levels = {h["text"]: h["level"] for h in res["headings"]}
        assert levels["Introduction"] == 1
        assert levels["Results"] == 2

    def test_outline_no_headings(self, simple_doc):
        res = reading.get_outline(simple_doc, json_mode=True)
        assert res["status"] == "WARNING"
        assert res["headings"] == []

    def test_outline_text_mode(self, heading_doc):
        res = reading.get_outline(heading_doc)
        assert "Introduction" in res
        assert "Methods" in res


class TestFullText:
    def test_basic_extraction(self, simple_doc):
        res = reading.full_text(simple_doc, json_mode=True)
        assert res["status"] == "SUCCESS"
        texts = [p["text"] for p in res["paragraphs"]]
        assert "First paragraph" in texts
        assert "Fifth paragraph" in texts

    def test_lxml_fast_path_matches_python_docx(self, simple_doc):
        """LXML fast-path (no formatting) must return same paragraph count as slow path."""
        fast = reading.full_text(simple_doc, include_formatting=False, json_mode=True)
        slow = reading.full_text(simple_doc, include_formatting=True, json_mode=True)
        # Both paths must see the same paragraphs
        assert len(fast["paragraphs"]) == len(slow["paragraphs"])
        # LXML path returns 'text'; formatting path returns 'runs' — check text field only
        fast_texts = [p["text"] for p in fast["paragraphs"]]
        for i, fast_text in enumerate(fast_texts):
            assert fast_text in slow["paragraphs"][i].get("runs", fast_text)

    def test_compact_skips_empty(self, simple_doc):
        # simple_doc has no empty paragraphs
        res = reading.full_text(simple_doc, compact=True, json_mode=True)
        for p in res["paragraphs"]:
            assert p["text"].strip() != ""

    def test_max_chars_truncates(self, simple_doc):
        res = reading.full_text(simple_doc, max_chars=5)
        assert "TRUNCATED" in res

    def test_include_formatting(self, multirun_doc):
        res = reading.full_text(multirun_doc, include_formatting=True, json_mode=True)
        assert res["status"] == "SUCCESS"
        # First para has runs with formatting info
        assert len(res["paragraphs"]) > 0


class TestReadRange:
    def test_valid_range(self, simple_doc):
        res = reading.read_range(simple_doc, 0, 2, json_mode=True)
        assert res["status"] == "SUCCESS"
        assert len(res["paragraphs"]) == 3

    def test_single_paragraph(self, simple_doc):
        res = reading.read_range(simple_doc, 1, 1, json_mode=True)
        assert res["status"] == "SUCCESS"
        assert res["paragraphs"][0]["text"] == "Second paragraph"

    def test_out_of_range_end(self, simple_doc):
        doc, _ = load_document(simple_doc)
        n = len(doc.paragraphs)
        res = reading.read_range(simple_doc, 0, n, json_mode=True)
        assert res["status"] == "ERROR"

    def test_negative_start(self, simple_doc):
        res = reading.read_range(simple_doc, -1, 2, json_mode=True)
        assert res["status"] == "ERROR"

    def test_last_valid_index(self, simple_doc):
        doc, _ = load_document(simple_doc)
        n = len(doc.paragraphs)
        res = reading.read_range(simple_doc, n - 1, n - 1, json_mode=True)
        assert res["status"] == "SUCCESS"


class TestSearchText:
    def test_finds_match(self, simple_doc):
        res = reading.search_text(simple_doc, "First", json_mode=True)
        assert res["status"] == "SUCCESS"
        assert res["match_count"] == 1

    def test_no_match(self, simple_doc):
        res = reading.search_text(simple_doc, "XNOTFOUND", json_mode=True)
        assert res["status"] == "SUCCESS"
        assert res["match_count"] == 0

    def test_regex_search(self, simple_doc):
        res = reading.search_text(simple_doc, r"[A-Z]\w+ paragraph", json_mode=True)
        assert res["match_count"] >= 1

    def test_case_insensitive(self, simple_doc):
        res_lower = reading.search_text(simple_doc, "first", json_mode=True)
        res_upper = reading.search_text(simple_doc, "FIRST", json_mode=True)
        assert res_lower["match_count"] == res_upper["match_count"]

    def test_context_lines(self, simple_doc):
        res = reading.search_text(simple_doc, "Third", context_lines=1, json_mode=True)
        assert res["match_count"] == 1
        match = res["matches"][0]
        assert len(match["context_before"]) <= 1
        assert len(match["context_after"]) <= 1
