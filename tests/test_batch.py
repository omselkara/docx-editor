"""Tests for docx_engine.batch_tools — batch execute, diff, backup/undo."""
from __future__ import annotations

import json
import os
import sys

import pytest

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../scripts")))

from docx_engine import batch_tools
from docx_engine.core import load_document


# ─── Backup / Undo ───────────────────────────────────────────────────────────

class TestBackupUndo:
    def test_create_backup(self, simple_doc):
        path = batch_tools.create_backup(simple_doc)
        assert path is not None
        assert os.path.exists(path)

    def test_undo_restores(self, simple_doc, tmp_path):
        from docx_engine import editing
        batch_tools.create_backup(simple_doc)
        editing.replace_text(simple_doc, "First", "MODIFIED")
        doc, _ = load_document(simple_doc)
        assert doc.paragraphs[0].text == "MODIFIED paragraph"

        res = batch_tools.undo(simple_doc)
        assert "SUCCESS" in res
        doc, _ = load_document(simple_doc)
        assert doc.paragraphs[0].text == "First paragraph"

    def test_undo_no_backup(self, tmp_path):
        path = str(tmp_path / "no_backup.docx")
        from docx_engine.core import create_document
        create_document(path)
        res = batch_tools.undo(path)
        assert res.startswith("ERROR")

    def test_rolling_backups(self, simple_doc):
        for _ in range(4):
            batch_tools.create_backup(simple_doc)
        # Should have .bak and up to .bak3 — but .bak4 must not exist
        assert not os.path.exists(simple_doc + ".bak4")


# ─── Diff Documents ──────────────────────────────────────────────────────────

class TestDiffDocuments:
    def test_identical_docs(self, simple_doc, tmp_path):
        import shutil
        copy = str(tmp_path / "copy.docx")
        shutil.copy2(simple_doc, copy)
        res = batch_tools.diff_documents(simple_doc, copy)
        assert "identical" in res.lower()

    def test_diff_detects_addition(self, simple_doc, tmp_path):
        import shutil
        from docx_engine import editing
        modified = str(tmp_path / "modified.docx")
        shutil.copy2(simple_doc, modified)
        editing.insert_paragraph(modified, "New paragraph")
        res = batch_tools.diff_documents(simple_doc, modified, json_mode=True)
        assert res["status"] == "SUCCESS"
        additions = [c for c in res["changes"] if c["type"] == "+"]
        assert len(additions) >= 1

    def test_diff_detects_deletion(self, simple_doc, tmp_path):
        import shutil
        from docx_engine import editing
        modified = str(tmp_path / "modified.docx")
        shutil.copy2(simple_doc, modified)
        editing.delete_paragraphs(modified, [0])
        res = batch_tools.diff_documents(simple_doc, modified, json_mode=True)
        deletions = [c for c in res["changes"] if c["type"] == "-"]
        assert len(deletions) >= 1

    def test_diff_json_mode(self, simple_doc, tmp_path):
        import shutil
        copy = str(tmp_path / "copy.docx")
        shutil.copy2(simple_doc, copy)
        res = batch_tools.diff_documents(simple_doc, copy, json_mode=True)
        assert "status" in res
        assert "changes" in res


# ─── Batch Execute ───────────────────────────────────────────────────────────

class TestBatchExecute:
    def _write_commands(self, tmp_path, commands: list[dict], on_error: str = "stop") -> str:
        path = str(tmp_path / "commands.json")
        with open(path, "w") as f:
            json.dump({"on_error": on_error, "commands": commands}, f)
        return path

    def test_successful_batch(self, simple_doc, tmp_path):
        cmds = self._write_commands(tmp_path, [
            {"action": "replace_text", "args": {"find_pattern": "First", "replace_with": "1st"}},
            {"action": "replace_text", "args": {"find_pattern": "Second", "replace_with": "2nd"}},
        ])
        res = batch_tools.batch_execute(simple_doc, cmds, json_mode=True)
        assert res["status"] == "SUCCESS"
        assert res["succeeded"] == 2
        assert res["failed"] == 0

    def test_rollback_on_error(self, simple_doc, tmp_path):
        original_text = load_document(simple_doc)[0].paragraphs[0].text
        cmds = self._write_commands(tmp_path, [
            {"action": "replace_text", "args": {"find_pattern": "First", "replace_with": "CHANGED"}},
            {"action": "INVALID_ACTION_XYZ", "args": {}},
        ])
        res = batch_tools.batch_execute(simple_doc, cmds, json_mode=True)
        assert res["aborted"] is True
        # Document should be rolled back
        doc, _ = load_document(simple_doc)
        assert doc.paragraphs[0].text == original_text

    def test_continue_on_error(self, simple_doc, tmp_path):
        cmds = self._write_commands(tmp_path, [
            {"action": "INVALID_ACTION", "args": {}},
            {"action": "replace_text", "args": {"find_pattern": "First", "replace_with": "1st"}},
        ], on_error="continue")
        res = batch_tools.batch_execute(simple_doc, cmds, json_mode=True)
        assert res["aborted"] is False
        assert res["succeeded"] >= 1

    def test_dry_run(self, simple_doc, tmp_path):
        cmds = self._write_commands(tmp_path, [
            {"action": "replace_text", "args": {"find_pattern": "First", "replace_with": "DRYRUN"}},
        ])
        res = batch_tools.batch_execute(simple_doc, cmds, dry_run=True, json_mode=True)
        assert res["dry_run"] is True
        # Document must be unchanged
        doc, _ = load_document(simple_doc)
        assert doc.paragraphs[0].text == "First paragraph"

    def test_empty_commands(self, simple_doc, tmp_path):
        cmds = self._write_commands(tmp_path, [])
        res = batch_tools.batch_execute(simple_doc, cmds)
        assert "WARNING" in res or "No commands" in res
