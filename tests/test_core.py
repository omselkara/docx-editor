"""Tests for docx_engine.core — create, load, save, info, metadata."""
from __future__ import annotations

import os
import sys

import pytest

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../scripts")))

from docx_engine import core


class TestCreateDocument:
    def test_creates_file(self, tmp_path):
        path = str(tmp_path / "new.docx")
        res = core.create_document(path)
        assert "SUCCESS" in res
        assert os.path.exists(path)

    def test_creates_valid_docx(self, tmp_path):
        path = str(tmp_path / "new.docx")
        core.create_document(path)
        doc, err = core.load_document(path)
        assert err is None
        assert doc is not None


class TestLoadDocument:
    def test_load_existing(self, simple_doc):
        doc, err = core.load_document(simple_doc)
        assert err is None
        assert doc is not None

    def test_load_missing(self):
        doc, err = core.load_document("/nonexistent/file.docx")
        assert doc is None
        assert err.startswith("ERROR")

    def test_load_not_a_docx(self, tmp_path):
        path = str(tmp_path / "notadocx.docx")
        with open(path, "w") as f:
            f.write("not a valid docx file")
        doc, err = core.load_document(path)
        assert doc is None
        assert err.startswith("ERROR")


class TestGetInfo:
    def test_returns_all_fields(self, simple_doc):
        res = core.get_info(simple_doc, json_mode=True)
        assert res["status"] == "SUCCESS"
        info = res["info"]
        assert "paragraph_count" in info
        assert "table_count" in info
        assert "word_count" in info
        assert "char_count" in info
        assert "section_count" in info

    def test_paragraph_count_correct(self, simple_doc):
        res = core.get_info(simple_doc, json_mode=True)
        assert res["info"]["paragraph_count"] == 5

    def test_text_mode(self, simple_doc):
        res = core.get_info(simple_doc)
        assert "DOCUMENT INFORMATION" in res

    def test_table_count(self, table_doc):
        res = core.get_info(table_doc, json_mode=True)
        assert res["info"]["table_count"] == 1

    def test_word_count_nonzero(self, simple_doc):
        res = core.get_info(simple_doc, json_mode=True)
        assert res["info"]["word_count"] > 0


class TestSetMetadata:
    def test_set_title(self, simple_doc):
        res = core.set_metadata(simple_doc, title="My Test Document")
        assert "SUCCESS" in res
        info = core.get_info(simple_doc, json_mode=True)
        assert info["info"]["title"] == "My Test Document"

    def test_set_author(self, simple_doc):
        res = core.set_metadata(simple_doc, author="Test Author")
        assert "SUCCESS" in res
        info = core.get_info(simple_doc, json_mode=True)
        assert info["info"]["author"] == "Test Author"

    def test_set_multiple_fields(self, simple_doc):
        res = core.set_metadata(simple_doc, title="Doc", subject="Testing", keywords="pytest")
        assert "SUCCESS" in res

    def test_no_fields_returns_warning(self, simple_doc):
        res = core.set_metadata(simple_doc)
        assert res.startswith("WARNING")
