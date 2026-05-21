"""Tests for footnote and endnote CRUD operations (update, delete)."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from docx_mcp import server


def _j(result: str) -> dict | list:
    return json.loads(result)


# ═══════════════════════════════════════════════════════════════════════════
#  TestFootnoteCRUD
# ═══════════════════════════════════════════════════════════════════════════


class TestFootnoteCRUD:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_update_footnote(self):
        """Update existing footnote #1 text and verify the result."""
        result = _j(server.update_footnote(1, "Updated footnote text."))
        assert result["footnote_id"] == 1
        assert result["text"] == "Updated footnote text."
        # Verify via get_footnotes
        footnotes = _j(server.get_footnotes())
        fn1 = next(f for f in footnotes if f["id"] == 1)
        assert "Updated footnote text." in fn1["text"]

    def test_update_footnote_not_found(self):
        """Updating a non-existent footnote raises ValueError."""
        with pytest.raises(ValueError, match="not found"):
            server.update_footnote(999, "Should fail")

    def test_update_footnote_builtin_rejected(self):
        """Updating built-in footnote (id < 1) raises ValueError."""
        with pytest.raises(ValueError):
            server.update_footnote(0, "Should fail")

    def test_delete_footnote_removes_from_xml(self):
        """delete_footnote removes the definition from footnotes.xml."""
        result = _j(server.delete_footnote(1))
        assert result["deleted"] == 1
        footnotes = _j(server.get_footnotes())
        ids = [f["id"] for f in footnotes]
        assert 1 not in ids

    def test_delete_footnote_removes_reference(self):
        """delete_footnote also removes the footnoteReference run in document.xml."""
        server.delete_footnote(1)
        # validate_footnotes should still report valid (no dangling refs)
        validation = _j(server.validate_footnotes())
        assert validation["valid"] is True
        assert 1 not in validation.get("missing_definitions", [])

    def test_delete_footnote_not_found(self):
        """Deleting a non-existent footnote raises ValueError."""
        with pytest.raises(ValueError, match="not found"):
            server.delete_footnote(999)

    def test_update_footnote_then_read_back(self):
        """Round-trip: add a new footnote, update it, confirm text changed."""
        add_result = _j(server.add_footnote("00000004", "Initial text"))
        fid = add_result["footnote_id"]
        _j(server.update_footnote(fid, "Revised text"))
        footnotes = _j(server.get_footnotes())
        fn = next(f for f in footnotes if f["id"] == fid)
        assert "Revised text" in fn["text"]

    def test_consecutive_footnotes_produce_comma_delimiter(self):
        """Two add_footnote() calls on the same paragraph insert a superscript comma between refs."""
        from lxml import etree

        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        W14 = "http://schemas.microsoft.com/office/word/2010/wordml"

        r1 = _j(server.add_footnote("00000004", "First source"))
        r2 = _j(server.add_footnote("00000004", "Second source"))
        id1, id2 = r1["footnote_id"], r2["footnote_id"]

        # Inspect raw document XML via the server's internal document registry
        doc_obj = server._docs[server._DEFAULT_HANDLE]
        doc_tree = doc_obj._tree("word/document.xml")

        # Find the paragraph by paraId
        para = None
        for p in doc_tree.iter(f"{{{W}}}p"):
            if p.get(f"{{{W14}}}paraId") == "00000004":
                para = p
                break
        assert para is not None, "Paragraph 00000004 not found"

        # Collect the tail of children: [... fn_ref(id1), comma_run, fn_ref(id2)]
        children = list(para)
        fn_runs = [
            c for c in children
            if c.tag == f"{{{W}}}r" and c.find(f"{{{W}}}footnoteReference") is not None
        ]
        assert len(fn_runs) >= 2, "Expected at least two footnote reference runs"

        # The runs for id1 and id2 should be separated by exactly one comma run
        idx1 = children.index(fn_runs[-2])
        idx2 = children.index(fn_runs[-1])
        assert idx2 == idx1 + 2, "Comma run should sit between the two consecutive fn ref runs"

        between = children[idx1 + 1]
        assert between.tag == f"{{{W}}}r"
        t_el = between.find(f"{{{W}}}t")
        assert t_el is not None and t_el.text == ",", "Separator run must contain a single comma"

        # Confirm the footnote ref IDs are correct
        ref1 = fn_runs[-2].find(f"{{{W}}}footnoteReference").get(f"{{{W}}}id")
        ref2 = fn_runs[-1].find(f"{{{W}}}footnoteReference").get(f"{{{W}}}id")
        assert ref1 == str(id1)
        assert ref2 == str(id2)


# ═══════════════════════════════════════════════════════════════════════════
#  TestEndnoteCRUD
# ═══════════════════════════════════════════════════════════════════════════


class TestEndnoteCRUD:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_update_endnote(self):
        """Update existing endnote #1 text and verify the result."""
        result = _j(server.update_endnote(1, "Updated endnote text."))
        assert result["endnote_id"] == 1
        assert result["text"] == "Updated endnote text."
        # Verify via get_endnotes
        endnotes = _j(server.get_endnotes())
        en1 = next(e for e in endnotes if e["id"] == 1)
        assert "Updated endnote text." in en1["text"]

    def test_update_endnote_not_found(self):
        """Updating a non-existent endnote raises ValueError."""
        with pytest.raises(ValueError, match="not found"):
            server.update_endnote(999, "Should fail")

    def test_update_endnote_builtin_rejected(self):
        """Updating built-in endnote (id < 1) raises ValueError."""
        with pytest.raises(ValueError):
            server.update_endnote(0, "Should fail")

    def test_delete_endnote_removes_from_xml(self):
        """delete_endnote removes the definition from endnotes.xml."""
        result = _j(server.delete_endnote(1))
        assert result["deleted"] == 1
        endnotes = _j(server.get_endnotes())
        ids = [e["id"] for e in endnotes]
        assert 1 not in ids

    def test_delete_endnote_removes_reference(self):
        """delete_endnote also removes the endnoteReference run in document.xml."""
        server.delete_endnote(1)
        # validate_endnotes should report valid (no dangling refs)
        validation = _j(server.validate_endnotes())
        assert validation["valid"] is True
        assert 1 not in validation.get("orphaned_refs", [])

    def test_delete_endnote_not_found(self):
        """Deleting a non-existent endnote raises ValueError."""
        with pytest.raises(ValueError, match="not found"):
            server.delete_endnote(999)

    def test_update_endnote_then_read_back(self):
        """Round-trip: add a new endnote, update it, confirm text changed."""
        add_result = _j(server.add_endnote("00000004", "Initial endnote"))
        eid = add_result["endnote_id"]
        _j(server.update_endnote(eid, "Revised endnote"))
        endnotes = _j(server.get_endnotes())
        en = next(e for e in endnotes if e["id"] == eid)
        assert "Revised endnote" in en["text"]
