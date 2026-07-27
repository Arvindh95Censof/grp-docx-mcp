"""Tests for LayoutMixin — paragraph enumeration, bulk fill, paraId injection,
style map extraction, and table styling."""

from __future__ import annotations

import uuid
from pathlib import Path

import pytest
from lxml import etree

from docx_mcp.document import W14, DocxDocument, W

# ── Helpers ──────────────────────────────────────────────────────────────────


def _make_doc(tmp_path: Path) -> DocxDocument:
    out = str(tmp_path / "test.docx")
    return DocxDocument.create(out)


def _inject_para(doc: DocxDocument, text: str = "placeholder", *, bold: bool = False) -> str:
    """Inject a plain paragraph into the body and return its para_id."""
    tree = doc._tree("word/document.xml")
    body = tree.find(f"{W}body")
    para_id = uuid.uuid4().hex[:8].upper()
    p = etree.Element(f"{W}p")
    p.set(f"{W14}paraId", para_id)
    r = etree.SubElement(p, f"{W}r")
    if bold:
        rpr = etree.SubElement(r, f"{W}rPr")
        etree.SubElement(rpr, f"{W}b")
    t = etree.SubElement(r, f"{W}t")
    t.text = text
    body.insert(len(body) - 1, p)
    doc._mark("word/document.xml")
    return para_id


def _strip_para_ids(doc: DocxDocument) -> None:
    tree = doc._tree("word/document.xml")
    for para in tree.iter(f"{W}p"):
        if para.get(f"{W14}paraId"):
            del para.attrib[f"{W14}paraId"]
    doc._mark("word/document.xml")


def _index_of(doc: DocxDocument, para_id: str) -> int:
    """Look up a paragraph's current index by its paraId.

    `_inject_para` always inserts before the document's trailing (originally
    default) paragraph, so a freshly-injected paragraph is never reliably
    "last" — callers must resolve position by paraId, not by list position.
    """
    return next(p["index"] for p in doc.get_all_paragraphs() if p["paraId"] == para_id)


# ── inject_para_ids ───────────────────────────────────────────────────────────


class TestInjectParaIds:
    def test_injects_missing_ids(self, tmp_path):
        doc = _make_doc(tmp_path)
        _inject_para(doc, "one")
        _inject_para(doc, "two")
        total_paras = len(doc.get_all_paragraphs())
        _strip_para_ids(doc)

        result = doc.inject_para_ids()
        assert result["injected"] == total_paras

        tree = doc._tree("word/document.xml")
        for para in tree.iter(f"{W}p"):
            assert para.get(f"{W14}paraId")

    def test_no_op_when_all_present(self, tmp_path):
        doc = _make_doc(tmp_path)
        _inject_para(doc, "one")
        result = doc.inject_para_ids()
        assert result["injected"] == 0


# ── get_all_paragraphs ────────────────────────────────────────────────────────


class TestGetAllParagraphs:
    def test_returns_index_style_text_paraid(self, tmp_path):
        doc = _make_doc(tmp_path)
        pid = _inject_para(doc, "hello world")
        paras = doc.get_all_paragraphs()
        assert any(p["text"] == "hello world" and p["paraId"] == pid for p in paras)
        assert all({"index", "style", "text", "paraId"} <= p.keys() for p in paras)

    def test_captures_paragraph_style(self, tmp_path):
        doc = _make_doc(tmp_path)
        doc.create_style("Heading9000", "paragraph")
        pid = _inject_para(doc, "styled")
        tree = doc._tree("word/document.xml")
        para = next(p for p in tree.iter(f"{W}p") if p.get(f"{W14}paraId") == pid)
        ppr = etree.SubElement(para, f"{W}pPr")
        pstyle = etree.SubElement(ppr, f"{W}pStyle")
        pstyle.set(f"{W}val", "Heading9000")
        para.insert(0, ppr)
        doc._mark("word/document.xml")

        entry = next(p for p in doc.get_all_paragraphs() if p["paraId"] == pid)
        assert entry["style"] == "Heading9000"


# ── fill_paragraph_by_index ───────────────────────────────────────────────────


class TestFillParagraphByIndex:
    def test_replaces_text(self, tmp_path):
        doc = _make_doc(tmp_path)
        pid = _inject_para(doc, "old text")
        idx = _index_of(doc, pid)

        result = doc.fill_paragraph_by_index(idx, "new text")
        assert result == {"index": idx, "text": "new text"}
        assert doc.get_all_paragraphs()[idx]["text"] == "new text"

    def test_preserves_run_formatting(self, tmp_path):
        doc = _make_doc(tmp_path)
        pid = _inject_para(doc, "bold text", bold=True)
        idx = _index_of(doc, pid)

        doc.fill_paragraph_by_index(idx, "still bold")
        tree = doc._tree("word/document.xml")
        body = tree.find(f"{W}body")
        para = list(body.iter(f"{W}p"))[idx]
        run = para.find(f"{W}r")
        assert run.find(f"{W}rPr").find(f"{W}b") is not None

    def test_out_of_range_raises(self, tmp_path):
        doc = _make_doc(tmp_path)
        _inject_para(doc, "only one")
        paras = doc.get_all_paragraphs()
        with pytest.raises(ValueError, match="out of range"):
            doc.fill_paragraph_by_index(len(paras) + 5, "x")

    def test_negative_index_raises(self, tmp_path):
        doc = _make_doc(tmp_path)
        _inject_para(doc, "only one")
        with pytest.raises(ValueError, match="out of range"):
            doc.fill_paragraph_by_index(-1, "x")


# ── bulk_replace_text ─────────────────────────────────────────────────────────


class TestBulkReplaceText:
    def test_replaces_matching_text(self, tmp_path):
        doc = _make_doc(tmp_path)
        _inject_para(doc, "Hello {{NAME}}, welcome to {{PLACE}}.")

        result = doc.bulk_replace_text({"{{NAME}}": "Alice", "{{PLACE}}": "Wonderland"})
        assert result["replacements"] == {"{{NAME}}": 1, "{{PLACE}}": 1}
        assert result["total"] == 2

        texts = [p["text"] for p in doc.get_all_paragraphs()]
        assert "Hello Alice, welcome to Wonderland." in texts

    def test_counts_multiple_occurrences(self, tmp_path):
        doc = _make_doc(tmp_path)
        _inject_para(doc, "foo foo foo")
        result = doc.bulk_replace_text({"foo": "bar"})
        assert result["replacements"]["foo"] == 3
        assert result["total"] == 3

    def test_no_match_leaves_document_unmodified(self, tmp_path):
        doc = _make_doc(tmp_path)
        _inject_para(doc, "unrelated content")
        result = doc.bulk_replace_text({"NOPE": "x"})
        assert result["replacements"] == {"NOPE": 0}
        assert result["total"] == 0

    def test_preserves_run_formatting(self, tmp_path):
        doc = _make_doc(tmp_path)
        _inject_para(doc, "{{X}}", bold=True)
        doc.bulk_replace_text({"{{X}}": "filled"})
        tree = doc._tree("word/document.xml")
        run = next(r for r in tree.iter(f"{W}r") if r.find(f"{W}t") is not None)
        assert run.find(f"{W}rPr").find(f"{W}b") is not None


# ── apply_table_cell_style ────────────────────────────────────────────────────


class TestApplyTableCellStyle:
    def test_applies_style_to_all_cells(self, tmp_path):
        doc = _make_doc(tmp_path)
        pid = _inject_para(doc, "before table")
        doc.add_table(pid, rows=2, cols=2)

        result = doc.apply_table_cell_style(0, "TableText")
        assert result["cells_updated"] == 4

        tree = doc._tree("word/document.xml")
        tbl = next(tree.iter(f"{W}tbl"))
        for tc in tbl.iter(f"{W}tc"):
            para = tc.find(f"{W}p")
            pstyle = para.find(f"{W}pPr").find(f"{W}pStyle")
            assert pstyle.get(f"{W}val") == "TableText"

    def test_replaces_existing_pstyle(self, tmp_path):
        doc = _make_doc(tmp_path)
        pid = _inject_para(doc, "before table")
        doc.add_table(pid, rows=1, cols=1)
        doc.apply_table_cell_style(0, "First")
        doc.apply_table_cell_style(0, "Second")

        tree = doc._tree("word/document.xml")
        tbl = next(tree.iter(f"{W}tbl"))
        tc = next(tbl.iter(f"{W}tc"))
        styles = tc.find(f"{W}p").find(f"{W}pPr").findall(f"{W}pStyle")
        assert len(styles) == 1
        assert styles[0].get(f"{W}val") == "Second"

    def test_out_of_range_raises(self, tmp_path):
        doc = _make_doc(tmp_path)
        with pytest.raises(IndexError, match="out of range"):
            doc.apply_table_cell_style(0, "X")


# ── copy_table_style ──────────────────────────────────────────────────────────


class TestCopyTableStyle:
    def test_copies_tblpr(self, tmp_path):
        doc = _make_doc(tmp_path)
        pid = _inject_para(doc, "p")
        doc.add_table(pid, rows=1, cols=1)
        doc.add_table(pid, rows=1, cols=1)

        tree = doc._tree("word/document.xml")
        src_tbl = list(tree.iter(f"{W}tbl"))[0]
        src_pr = src_tbl.find(f"{W}tblPr")
        border = etree.SubElement(src_pr, f"{W}tblBorders")
        etree.SubElement(border, f"{W}top").set(f"{W}val", "single")
        doc._mark("word/document.xml")

        result = doc.copy_table_style(0, 1)
        assert result == {"source": 0, "target": 1, "style_copied": True}

        tgt_tbl = list(doc._tree("word/document.xml").iter(f"{W}tbl"))[1]
        assert tgt_tbl.find(f"{W}tblPr").find(f"{W}tblBorders") is not None

    def test_preserves_target_alignment(self, tmp_path):
        doc = _make_doc(tmp_path)
        pid = _inject_para(doc, "p")
        doc.add_table(pid, rows=1, cols=1)
        doc.add_table(pid, rows=1, cols=1)

        tree = doc._tree("word/document.xml")
        tgt_tbl = list(tree.iter(f"{W}tbl"))[1]
        jc = etree.SubElement(tgt_tbl.find(f"{W}tblPr"), f"{W}jc")
        jc.set(f"{W}val", "center")
        doc._mark("word/document.xml")

        doc.copy_table_style(0, 1)
        tgt_tbl = list(doc._tree("word/document.xml").iter(f"{W}tbl"))[1]
        assert tgt_tbl.find(f"{W}tblPr").find(f"{W}jc").get(f"{W}val") == "center"

    def test_source_without_tblpr_raises(self, tmp_path):
        doc = _make_doc(tmp_path)
        pid = _inject_para(doc, "p")
        doc.add_table(pid, rows=1, cols=1)
        doc.add_table(pid, rows=1, cols=1)

        tree = doc._tree("word/document.xml")
        src_tbl = list(tree.iter(f"{W}tbl"))[0]
        src_tbl.remove(src_tbl.find(f"{W}tblPr"))
        doc._mark("word/document.xml")

        with pytest.raises(ValueError, match="no tblPr"):
            doc.copy_table_style(0, 1)

    def test_source_out_of_range_raises(self, tmp_path):
        doc = _make_doc(tmp_path)
        pid = _inject_para(doc, "p")
        doc.add_table(pid, rows=1, cols=1)
        with pytest.raises(IndexError, match="Source table"):
            doc.copy_table_style(5, 0)

    def test_target_out_of_range_raises(self, tmp_path):
        doc = _make_doc(tmp_path)
        pid = _inject_para(doc, "p")
        doc.add_table(pid, rows=1, cols=1)
        with pytest.raises(IndexError, match="Target table"):
            doc.copy_table_style(0, 5)


# ── extract_template_structure ────────────────────────────────────────────────


class TestExtractTemplateStructure:
    def test_returns_paragraphs_and_tables(self, tmp_path):
        doc = _make_doc(tmp_path)
        pid = _inject_para(doc, "hello")
        doc.add_table(pid, rows=2, cols=3)

        structure = doc.extract_template_structure()
        assert any(p["text"] == "hello" for p in structure["paragraphs"])
        assert structure["tables"][0]["row_count"] == 2
        assert structure["tables"][0]["col_count"] == 3

    def test_truncates_long_paragraph_text(self, tmp_path):
        doc = _make_doc(tmp_path)
        pid = _inject_para(doc, "x" * 300)
        idx = _index_of(doc, pid)
        structure = doc.extract_template_structure()
        assert len(structure["paragraphs"][idx]["text"]) == 200

    def test_truncates_long_cell_text(self, tmp_path):
        doc = _make_doc(tmp_path)
        pid = _inject_para(doc, "p")
        doc.add_table(pid, rows=1, cols=1)
        tree = doc._tree("word/document.xml")
        tc = next(tree.iter(f"{W}tc"))
        p = tc.find(f"{W}p")
        r = etree.SubElement(p, f"{W}r")
        t = etree.SubElement(r, f"{W}t")
        t.text = "y" * 300
        doc._mark("word/document.xml")

        structure = doc.extract_template_structure()
        assert len(structure["tables"][0]["cells"][0][0]) == 100

    def test_empty_table_has_zero_col_count(self, tmp_path):
        doc = _make_doc(tmp_path)
        pid = _inject_para(doc, "p")
        doc.add_table(pid, rows=0, cols=2)
        structure = doc.extract_template_structure()
        assert structure["tables"][0]["col_count"] == 0


# ── extract_style_map ─────────────────────────────────────────────────────────


class TestExtractStyleMap:
    def test_page_layout(self, tmp_path):
        doc = _make_doc(tmp_path)
        tree = doc._tree("word/document.xml")
        body = tree.find(f"{W}body")
        sect_pr = body.find(f"{W}sectPr")
        if sect_pr is None:
            sect_pr = etree.SubElement(body, f"{W}sectPr")
        pg_sz = etree.SubElement(sect_pr, f"{W}pgSz")
        pg_sz.set(f"{W}w", "12240")
        pg_sz.set(f"{W}h", "15840")
        pg_sz.set(f"{W}orient", "landscape")
        pg_mar = etree.SubElement(sect_pr, f"{W}pgMar")
        for k, v in {
            "top": "1440",
            "bottom": "1440",
            "left": "1800",
            "right": "1800",
            "header": "720",
            "footer": "720",
            "gutter": "0",
        }.items():
            pg_mar.set(f"{W}{k}", v)
        cols = etree.SubElement(sect_pr, f"{W}cols")
        cols.set(f"{W}num", "2")
        cols.set(f"{W}space", "720")
        doc._mark("word/document.xml")

        style_map = doc.extract_style_map()
        assert style_map["page"]["width_dxa"] == "12240"
        assert style_map["page"]["orientation"] == "landscape"
        assert style_map["page"]["margins"]["left"] == "1800"
        assert "gutter" not in style_map["page"]["margins"] or style_map["page"]["margins"].get(
            "gutter"
        ) in ("0", None)
        assert style_map["page"]["columns"] == "2"
        assert style_map["page"]["col_space"] == "720"

    def test_doc_defaults(self, tmp_path):
        doc = _make_doc(tmp_path)
        styles = doc._tree("word/styles.xml")
        defaults = etree.SubElement(styles, f"{W}docDefaults")
        rpr_default = etree.SubElement(defaults, f"{W}rPrDefault")
        rpr = etree.SubElement(rpr_default, f"{W}rPr")
        sz = etree.SubElement(rpr, f"{W}sz")
        sz.set(f"{W}val", "24")
        ppr_default = etree.SubElement(defaults, f"{W}pPrDefault")
        ppr = etree.SubElement(ppr_default, f"{W}pPr")
        jc = etree.SubElement(ppr, f"{W}jc")
        jc.set(f"{W}val", "both")
        doc._mark("word/styles.xml")

        style_map = doc.extract_style_map()
        assert style_map["default"]["size_pt"] == 12
        assert style_map["default"]["alignment"] == "both"

    def test_paragraph_and_character_styles(self, tmp_path):
        doc = _make_doc(tmp_path)
        doc.create_style("Heading9000", "paragraph", based_on="Normal", next_style="Normal")
        doc.create_style("Emphasis9000", "character")

        styles = doc._tree("word/styles.xml")
        heading = next(
            s
            for s in styles.findall(f"{W}style")
            if s.find(f"{W}name").get(f"{W}val") == "Heading9000"
        )
        rpr = etree.SubElement(heading, f"{W}rPr")
        rfonts = etree.SubElement(rpr, f"{W}rFonts")
        rfonts.set(f"{W}ascii", "Calibri")
        rfonts.set(f"{W}cs", "Calibri-CS")
        sz = etree.SubElement(rpr, f"{W}sz")
        sz.set(f"{W}val", "32")
        szcs = etree.SubElement(rpr, f"{W}szCs")
        szcs.set(f"{W}val", "32")
        etree.SubElement(rpr, f"{W}b")
        etree.SubElement(rpr, f"{W}i")
        color = etree.SubElement(rpr, f"{W}color")
        color.set(f"{W}val", "FF0000")
        u = etree.SubElement(rpr, f"{W}u")
        u.set(f"{W}val", "single")
        etree.SubElement(rpr, f"{W}strike")
        kern = etree.SubElement(rpr, f"{W}kern")
        kern.set(f"{W}val", "16")
        spacing = etree.SubElement(rpr, f"{W}spacing")
        spacing.set(f"{W}val", "20")
        highlight = etree.SubElement(rpr, f"{W}highlight")
        highlight.set(f"{W}val", "yellow")

        ppr = etree.SubElement(heading, f"{W}pPr")
        pjc = etree.SubElement(ppr, f"{W}jc")
        pjc.set(f"{W}val", "center")
        pspacing = etree.SubElement(ppr, f"{W}spacing")
        pspacing.set(f"{W}before", "240")
        pspacing.set(f"{W}after", "120")
        pspacing.set(f"{W}line", "360")
        pspacing.set(f"{W}lineRule", "auto")
        pspacing.set(f"{W}beforeLines", "1")
        pspacing.set(f"{W}afterLines", "1")
        ind = etree.SubElement(ppr, f"{W}ind")
        ind.set(f"{W}left", "720")
        ind.set(f"{W}right", "720")
        ind.set(f"{W}firstLine", "240")
        ind.set(f"{W}hanging", "0")
        ind.set(f"{W}leftChars", "0")
        ind.set(f"{W}rightChars", "0")
        shd = etree.SubElement(ppr, f"{W}shd")
        shd.set(f"{W}fill", "CCCCCC")
        shd.set(f"{W}val", "clear")
        etree.SubElement(ppr, f"{W}keepNext")
        etree.SubElement(ppr, f"{W}pageBreakBefore")
        outline = etree.SubElement(ppr, f"{W}outlineLvl")
        outline.set(f"{W}val", "1")
        pbdr = etree.SubElement(ppr, f"{W}pBdr")
        for side in ("top", "bottom", "left", "right", "between", "bar"):
            b = etree.SubElement(pbdr, f"{W}{side}")
            b.set(f"{W}val", "single")
            b.set(f"{W}sz", "4")
            b.set(f"{W}color", "000000")
        inner_rpr = etree.SubElement(ppr, f"{W}rPr")
        inner_sz = etree.SubElement(inner_rpr, f"{W}sz")
        inner_sz.set(f"{W}val", "40")
        doc._mark("word/styles.xml")

        style_map = doc.extract_style_map()
        entry = style_map["para_styles"]["Heading9000"]
        assert entry["basedOn"] == "Normal"
        assert entry["next"] == "Normal"
        assert entry["font"] == "Calibri"
        assert entry["font_cs"] == "Calibri-CS"
        assert entry["bold"] is True
        assert entry["italic"] is True
        assert entry["color"] == "FF0000"
        assert entry["underline"] == "single"
        assert entry["strikethrough"] is True
        assert entry["kern"] == "16"
        assert entry["char_spacing"] == "20"
        assert entry["highlight"] == "yellow"
        assert entry["alignment"] == "center"
        assert entry["spacing_before"] == "240"
        assert entry["indent_left"] == "720"
        assert entry["shading_fill"] == "CCCCCC"
        assert entry["keep_next"] is True
        assert entry["page_break_before"] is True
        assert entry["outline_level"] == 1
        assert entry["borders"]["top"]["val"] == "single"
        # inner rPr (from pPrChange-style nesting) overrides size_pt
        assert entry["size_pt"] == 20
        assert "Emphasis9000" in style_map["char_styles"]

    def test_normal_style_fills_default(self, tmp_path):
        doc = _make_doc(tmp_path)
        styles = doc._tree("word/styles.xml")
        normal = next(
            (
                s
                for s in styles.findall(f"{W}style")
                if s.find(f"{W}name") is not None
                and s.find(f"{W}name").get(f"{W}val") == "Normal"
            ),
            None,
        )
        if normal is None:
            normal = etree.SubElement(styles, f"{W}style")
            normal.set(f"{W}type", "paragraph")
            normal.set(f"{W}styleId", "Normal")
            name_el = etree.SubElement(normal, f"{W}name")
            name_el.set(f"{W}val", "Normal")
        rpr = etree.SubElement(normal, f"{W}rPr")
        sz = etree.SubElement(rpr, f"{W}sz")
        sz.set(f"{W}val", "22")
        doc._mark("word/styles.xml")

        style_map = doc.extract_style_map()
        assert style_map["default"]["size_pt"] == 11

    def test_table_style_extraction(self, tmp_path):
        doc = _make_doc(tmp_path)
        pid = _inject_para(doc, "p")
        doc.add_table(pid, rows=2, cols=2)

        tree = doc._tree("word/document.xml")
        tbl = next(tree.iter(f"{W}tbl"))
        tbl_pr = tbl.find(f"{W}tblPr")
        jc = etree.SubElement(tbl_pr, f"{W}jc")
        jc.set(f"{W}val", "center")
        # add_table already created a tblW (w=0, type=auto) — mutate it in place
        # rather than appending a duplicate, since extract_style_map reads the
        # first tblW it finds.
        tw = tbl_pr.find(f"{W}tblW")
        tw.set(f"{W}w", "5000")
        tw.set(f"{W}type", "pct")
        borders = etree.SubElement(tbl_pr, f"{W}tblBorders")
        for side in ("top", "bottom", "left", "right", "insideH", "insideV"):
            b = etree.SubElement(borders, f"{W}{side}")
            b.set(f"{W}val", "single")
        cellmar = etree.SubElement(tbl_pr, f"{W}tblCellMar")
        for side in ("top", "bottom", "left", "right"):
            m = etree.SubElement(cellmar, f"{W}{side}")
            m.set(f"{W}w", "108")
            m.set(f"{W}type", "dxa")
        look = etree.SubElement(tbl_pr, f"{W}tblLook")
        look.set(f"{W}val", "04A0")
        look.set(f"{W}firstRow", "1")
        look.set(f"{W}lastRow", "0")
        look.set(f"{W}firstColumn", "1")
        look.set(f"{W}lastColumn", "0")
        look.set(f"{W}noHBand", "0")
        look.set(f"{W}noVBand", "1")

        first_row = tbl.findall(f"{W}tr")[0]
        for tc in first_row.findall(f"{W}tc"):
            tc_pr = tc.find(f"{W}tcPr")
            if tc_pr is None:
                tc_pr = etree.SubElement(tc, f"{W}tcPr")
                tc.insert(0, tc_pr)
            tcw = etree.SubElement(tc_pr, f"{W}tcW")
            tcw.set(f"{W}w", "2500")
            tcw.set(f"{W}type", "dxa")
            tcbdr = etree.SubElement(tc_pr, f"{W}tcBorders")
            for side in ("top", "bottom", "left", "right", "insideH", "insideV", "tl2br", "tr2bl"):
                b = etree.SubElement(tcbdr, f"{W}{side}")
                b.set(f"{W}val", "single")
            shd = etree.SubElement(tc_pr, f"{W}shd")
            shd.set(f"{W}fill", "DDDDDD")
            valign = etree.SubElement(tc_pr, f"{W}vAlign")
            valign.set(f"{W}val", "center")
        doc._mark("word/document.xml")

        style_map = doc.extract_style_map()
        tinfo = style_map["tables"][0]
        assert tinfo["style"] == "TableGrid"
        assert tinfo["alignment"] == "center"
        assert tinfo["width"] == {"value": "5000", "type": "pct"}
        assert tinfo["borders"]["top"]["val"] == "single"
        assert tinfo["cell_margins"]["top"] == {"w": "108", "type": "dxa"}
        assert tinfo["tbl_look"]["firstRow"] == "1"
        assert len(tinfo["col_widths_dxa"]) == 2
        assert tinfo["row_count"] == 2
        assert tinfo["first_row_cells"][0]["width"] == {"value": "2500", "type": "dxa"}
        assert tinfo["first_row_cells"][0]["borders"]["tl2br"]["val"] == "single"
        assert tinfo["first_row_cells"][0]["shading"] == "DDDDDD"
        assert tinfo["first_row_cells"][0]["v_align"] == "center"
        assert style_map["table_styles"] == ["TableGrid"]
        assert style_map["first_table_tblPr"] is not None

    def test_no_tables_or_extra_styles(self, tmp_path):
        doc = _make_doc(tmp_path)
        style_map = doc.extract_style_map()
        assert style_map["tables"] == []
        assert style_map["table_styles"] == []
        assert style_map["first_table_tblPr"] is None
