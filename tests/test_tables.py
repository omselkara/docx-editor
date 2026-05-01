"""Tests for docx_engine.tables — CRUD, merge, bounds checking."""
from __future__ import annotations

import os
import sys

import pytest

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../scripts")))

from docx_engine import tables
from docx_engine.core import load_document


class TestListTables:
    def test_with_table(self, table_doc):
        res = tables.list_tables(table_doc, json_mode=True)
        assert res["status"] == "SUCCESS"
        assert len(res["tables"]) == 1

    def test_no_tables(self, simple_doc):
        res = tables.list_tables(simple_doc, json_mode=True)
        assert res["status"] == "WARNING"
        assert res["tables"] == []


class TestReadTable:
    def test_basic_read(self, table_doc):
        res = tables.read_table(table_doc, 0, json_mode=True)
        assert res["status"] == "SUCCESS"
        assert res["rows"] == 3
        assert res["cols"] == 3

    def test_header_row(self, table_doc):
        res = tables.read_table(table_doc, 0, json_mode=True)
        headers = res["data"][0]
        assert "Name" in headers
        assert "Age" in headers
        assert "City" in headers

    def test_invalid_index(self, table_doc):
        res = tables.read_table(table_doc, 99, json_mode=True)
        assert res["status"] == "ERROR"

    def test_text_mode(self, table_doc):
        res = tables.read_table(table_doc, 0)
        assert "Name" in res


class TestCreateTable:
    def test_create_basic(self, blank_doc):
        res = tables.create_table(blank_doc, headers="A,B,C")
        assert "SUCCESS" in res
        doc, _ = load_document(blank_doc)
        assert len(doc.tables) == 1
        assert len(doc.tables[0].columns) == 3

    def test_create_with_data(self, blank_doc):
        res = tables.create_table(blank_doc, headers="Name,Age", rows_data="Alice,30;Bob,25")
        assert "SUCCESS" in res
        doc, _ = load_document(blank_doc)
        t = doc.tables[0]
        assert len(t.rows) == 3  # header + 2 data rows
        assert t.rows[1].cells[0].text == "Alice"
        assert t.rows[2].cells[0].text == "Bob"


class TestModifyCell:
    def test_modify_cell(self, table_doc):
        res = tables.modify_cell(table_doc, 0, 1, 0, "Charlie")
        assert "SUCCESS" in res
        doc, _ = load_document(table_doc)
        assert doc.tables[0].rows[1].cells[0].text == "Charlie"

    def test_invalid_table(self, table_doc):
        res = tables.modify_cell(table_doc, 99, 0, 0, "X")
        assert res.startswith("ERROR")

    def test_invalid_cell(self, table_doc):
        res = tables.modify_cell(table_doc, 0, 99, 99, "X")
        assert res.startswith("ERROR")


class TestAddDeleteRows:
    def test_add_row(self, table_doc):
        doc, _ = load_document(table_doc)
        orig_rows = len(doc.tables[0].rows)
        res = tables.add_row(table_doc, 0)
        assert "SUCCESS" in res
        doc, _ = load_document(table_doc)
        assert len(doc.tables[0].rows) == orig_rows + 1

    def test_add_row_with_values(self, table_doc):
        res = tables.add_row(table_doc, 0, values="Dave,35,Paris")
        assert "SUCCESS" in res
        doc, _ = load_document(table_doc)
        last_row = doc.tables[0].rows[-1]
        assert last_row.cells[0].text == "Dave"

    def test_delete_row(self, table_doc):
        doc, _ = load_document(table_doc)
        orig_rows = len(doc.tables[0].rows)
        res = tables.delete_row(table_doc, 0, 1)
        assert "SUCCESS" in res
        doc, _ = load_document(table_doc)
        assert len(doc.tables[0].rows) == orig_rows - 1

    def test_delete_row_invalid(self, table_doc):
        res = tables.delete_row(table_doc, 0, 999)
        assert res.startswith("ERROR")

    def test_delete_column(self, table_doc):
        doc, _ = load_document(table_doc)
        orig_cell_count = len(doc.tables[0].rows[0].cells)
        res = tables.delete_column(table_doc, 0, 0)
        assert "SUCCESS" in res
        doc, _ = load_document(table_doc)
        # After deletion each row has one fewer cell
        assert len(doc.tables[0].rows[0].cells) == orig_cell_count - 1


class TestMergeCells:
    def test_merge_cells(self, table_doc):
        res = tables.merge_cells(table_doc, 0, 0, 0, 0, 1)
        assert "SUCCESS" in res

    def test_invalid_table(self, table_doc):
        res = tables.merge_cells(table_doc, 99, 0, 0, 1, 1)
        assert res.startswith("ERROR")
