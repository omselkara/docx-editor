import os
import sys
import pytest

# Add scripts directory to path to allow importing docx_engine
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../scripts')))

from docx_engine import core
from docx_engine import reading

def test_document_creation(tmp_path):
    """Test that we can create a basic DOCX document."""
    doc_path = str(tmp_path / "test.docx")
    
    # Create
    res = core.create_document(doc_path)
    assert "SUCCESS" in res
    assert os.path.exists(doc_path)
    
    # Get Info
    info = core.get_info(doc_path, json_mode=True)
    assert info["status"] == "SUCCESS"
    # python-docx creates one empty paragraph by default
    assert info["info"]["paragraph_count"] >= 0

def test_document_reading(tmp_path):
    """Test reading basic properties."""
    doc_path = str(tmp_path / "test2.docx")
    core.create_document(doc_path)
    
    # Outline should be empty/warning
    outline = reading.get_outline(doc_path, json_mode=True)
    assert outline["status"] in ["WARNING", "SUCCESS"]
    
    # Full text fast path should work
    text = reading.full_text(doc_path, json_mode=True)
    assert text["status"] == "SUCCESS"
    assert text["paragraphs"] == []
