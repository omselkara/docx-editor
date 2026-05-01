"""Tests for docx_engine.formatting — format_text, format_paragraph, styles."""
from __future__ import annotations

import os
import sys

import pytest

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../scripts")))

from docx_engine import formatting
from docx_engine.core import load_document


class TestFormatText:
    def test_bold_by_index(self, simple_doc):
        res = formatting.format_text(simple_doc, para_indices=[0], bold=True)
        assert "SUCCESS" in res
        doc, _ = load_document(simple_doc)
        assert all(run.bold for run in doc.paragraphs[0].runs)

    def test_italic_by_index(self, simple_doc):
        res = formatting.format_text(simple_doc, para_indices=[0], italic=True)
        assert "SUCCESS" in res
        doc, _ = load_document(simple_doc)
        assert all(run.italic for run in doc.paragraphs[0].runs)

    def test_underline_by_index(self, simple_doc):
        res = formatting.format_text(simple_doc, para_indices=[0], underline=True)
        assert "SUCCESS" in res
        doc, _ = load_document(simple_doc)
        assert all(run.underline for run in doc.paragraphs[0].runs)

    def test_font_size(self, simple_doc):
        res = formatting.format_text(simple_doc, para_indices=[0], font_size=14)
        assert "SUCCESS" in res

    def test_font_color(self, simple_doc):
        res = formatting.format_text(simple_doc, para_indices=[0], font_color="#FF0000")
        assert "SUCCESS" in res

    def test_format_by_match(self, simple_doc):
        res = formatting.format_text(simple_doc, match="First", bold=True)
        assert "SUCCESS" in res

    def test_no_match_or_indices_returns_error(self, simple_doc):
        res = formatting.format_text(simple_doc, bold=True)
        assert res.startswith("ERROR")

    def test_match_not_found_returns_warning(self, simple_doc):
        res = formatting.format_text(simple_doc, match="XNOTFOUND", bold=True)
        assert res.startswith("WARNING")

    def test_multiple_indices(self, simple_doc):
        res = formatting.format_text(simple_doc, para_indices=[0, 1, 2], bold=True)
        assert "SUCCESS" in res


class TestFormatParagraph:
    def test_alignment_center(self, simple_doc):
        res = formatting.format_paragraph(simple_doc, para_indices=[0], alignment="center")
        assert "SUCCESS" in res

    def test_alignment_justify(self, simple_doc):
        res = formatting.format_paragraph(simple_doc, para_indices=[0], alignment="justify")
        assert "SUCCESS" in res

    def test_space_before_after(self, simple_doc):
        res = formatting.format_paragraph(simple_doc, para_indices=[0], space_before=6, space_after=6)
        assert "SUCCESS" in res

    def test_invalid_index_skipped(self, simple_doc):
        res = formatting.format_paragraph(simple_doc, para_indices=[999], alignment="center")
        assert res.startswith("WARNING")


class TestListStyles:
    def test_returns_styles(self, simple_doc):
        res = formatting.list_styles(simple_doc, json_mode=True)
        assert res["status"] == "SUCCESS"
        assert len(res["styles"]) > 0

    def test_filter_paragraph_styles(self, simple_doc):
        res = formatting.list_styles(simple_doc, style_type="paragraph", json_mode=True)
        assert res["status"] == "SUCCESS"

    def test_text_mode(self, simple_doc):
        res = formatting.list_styles(simple_doc)
        assert "STYLES" in res


class TestApplyStyle:
    def test_apply_normal_style(self, heading_doc):
        res = formatting.apply_style(heading_doc, para_indices=[1], style_name="Normal")
        assert "SUCCESS" in res

    def test_apply_nonexistent_style(self, simple_doc):
        res = formatting.apply_style(simple_doc, para_indices=[0], style_name="NonExistentStyle12345")
        assert res.startswith("ERROR")
