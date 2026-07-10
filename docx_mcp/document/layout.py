"""Layout mixin: paragraph enumeration, bulk fill, paraId injection, style map, table styling."""

from __future__ import annotations

import copy

from lxml import etree

from .base import W, W14


class LayoutMixin:
    """Tools for filling document templates without relying on paraIds."""

    def inject_para_ids(self) -> dict:
        """Assign w14:paraId to every paragraph that is missing one.

        Returns the count of paraIds injected. Run this once on templates
        that were created without paraIds so all other tools can target them.
        """
        doc = self._require("word/document.xml")
        injected = 0
        for para in doc.iter(f"{W}p"):
            if not para.get(f"{W14}paraId"):
                para.set(f"{W14}paraId", self._new_para_id())
                injected += 1
        if injected:
            self._mark("word/document.xml")
        return {"injected": injected}

    def get_all_paragraphs(self) -> list[dict]:
        """Return all body paragraphs with index, style, text, and paraId.

        Use this to discover paragraph indices for fill_paragraph_by_index,
        or paraIds (after inject_para_ids) for other tools.
        """
        doc = self._require("word/document.xml")
        body = doc.find(f"{W}body")
        result = []
        for idx, para in enumerate(body.iter(f"{W}p")):
            ppr = para.find(f"{W}pPr")
            style = ""
            if ppr is not None:
                ps = ppr.find(f"{W}pStyle")
                if ps is not None:
                    style = ps.get(f"{W}val", "")
            result.append(
                {
                    "index": idx,
                    "style": style,
                    "text": self._text(para),
                    "paraId": para.get(f"{W14}paraId", ""),
                }
            )
        return result

    def fill_paragraph_by_index(self, index: int, text: str) -> dict:
        """Replace all text in a paragraph by its zero-based index.

        Clears existing runs and inserts a single run with the new text,
        preserving the paragraph style and the first run's character formatting.

        Args:
            index: Zero-based paragraph index from get_all_paragraphs.
            text: New text content.
        """
        doc = self._require("word/document.xml")
        body = doc.find(f"{W}body")
        paras = list(body.iter(f"{W}p"))
        if index < 0 or index >= len(paras):
            raise ValueError(f"Paragraph index {index} out of range (0–{len(paras) - 1})")
        para = paras[index]

        # Capture rPr from first run to preserve character formatting
        rpr_bytes: bytes | None = None
        first_run = para.find(f"{W}r")
        if first_run is not None:
            rpr = first_run.find(f"{W}rPr")
            if rpr is not None:
                rpr_bytes = etree.tostring(rpr)

        # Remove all existing runs
        for run in para.findall(f"{W}r"):
            para.remove(run)

        # Insert single run with new text
        para.append(self._make_run(text, rpr_bytes))
        self._mark("word/document.xml")
        return {"index": index, "text": text}

    def bulk_replace_text(self, replacements: dict) -> dict:
        """Find and replace text across all body paragraphs.

        Handles text split across multiple runs by reconstructing the
        full paragraph text before matching.

        Args:
            replacements: Dict mapping search_text -> replacement_text.
                          Matching is case-sensitive.

        Returns:
            Dict with per-key replacement counts and a total.
        """
        doc = self._require("word/document.xml")
        body = doc.find(f"{W}body")
        counts: dict[str, int] = {k: 0 for k in replacements}
        modified = False

        for para in body.iter(f"{W}p"):
            full_text = self._text(para)
            new_text = full_text
            matched = False

            for search, replacement in replacements.items():
                if search in new_text:
                    new_text = new_text.replace(search, replacement)
                    counts[search] += full_text.count(search)
                    matched = True

            if matched:
                # Capture rPr from first run
                rpr_bytes: bytes | None = None
                first_run = para.find(f"{W}r")
                if first_run is not None:
                    rpr = first_run.find(f"{W}rPr")
                    if rpr is not None:
                        rpr_bytes = etree.tostring(rpr)

                # Remove all runs and rewrite as single run
                for run in para.findall(f"{W}r"):
                    para.remove(run)
                para.append(self._make_run(new_text, rpr_bytes))
                modified = True

        if modified:
            self._mark("word/document.xml")

        return {"replacements": counts, "total": sum(counts.values())}


    @staticmethod
    def _border_info(el: etree._Element, tag: str) -> dict:
        b = el.find(tag)
        if b is None:
            return {}
        return {
            k: b.get(f"{W}{k}", "")
            for k in ("val", "sz", "color", "space", "themeColor")
            if b.get(f"{W}{k}")
        }

    @staticmethod
    def _half_pt(el: etree._Element, tag: str) -> int | None:
        e = el.find(tag)
        if e is not None:
            v = e.get(f"{W}val", "")
            if v.isdigit():
                return int(v) // 2
        return None

    @staticmethod
    def _dxa(el: etree._Element, tag: str, attr: str = "w") -> str:
        e = el.find(tag)
        return e.get(f"{W}{attr}", "") if e is not None else ""

    def _parse_rpr(self, rpr: etree._Element) -> dict:
        """Extract all character properties from a w:rPr element."""
        entry: dict = {}
        if rpr is None:
            return entry
        rfonts = rpr.find(f"{W}rFonts")
        if rfonts is not None:
            entry["font"] = (
                rfonts.get(f"{W}ascii")
                or rfonts.get(f"{W}hAnsi")
                or rfonts.get(f"{W}cs")
                or ""
            )
            if rfonts.get(f"{W}cs"):
                entry["font_cs"] = rfonts.get(f"{W}cs")
        sz = rpr.find(f"{W}sz")
        if sz is not None and sz.get(f"{W}val", "").isdigit():
            entry["size_pt"] = int(sz.get(f"{W}val")) // 2
        szCs = rpr.find(f"{W}szCs")
        if szCs is not None and szCs.get(f"{W}val", "").isdigit():
            entry["size_cs_pt"] = int(szCs.get(f"{W}val")) // 2
        if rpr.find(f"{W}b") is not None:
            entry["bold"] = True
        if rpr.find(f"{W}i") is not None:
            entry["italic"] = True
        color = rpr.find(f"{W}color")
        if color is not None and color.get(f"{W}val"):
            entry["color"] = color.get(f"{W}val")
        u = rpr.find(f"{W}u")
        if u is not None and u.get(f"{W}val"):
            entry["underline"] = u.get(f"{W}val")
        if rpr.find(f"{W}strike") is not None:
            entry["strikethrough"] = True
        kern = rpr.find(f"{W}kern")
        if kern is not None and kern.get(f"{W}val"):
            entry["kern"] = kern.get(f"{W}val")
        spacing = rpr.find(f"{W}spacing")
        if spacing is not None and spacing.get(f"{W}val"):
            entry["char_spacing"] = spacing.get(f"{W}val")
        highlight = rpr.find(f"{W}highlight")
        if highlight is not None and highlight.get(f"{W}val"):
            entry["highlight"] = highlight.get(f"{W}val")
        return entry

    def _parse_ppr(self, ppr: etree._Element) -> dict:
        """Extract all paragraph properties from a w:pPr element."""
        entry: dict = {}
        if ppr is None:
            return entry
        jc = ppr.find(f"{W}jc")
        if jc is not None and jc.get(f"{W}val"):
            entry["alignment"] = jc.get(f"{W}val")
        spacing = ppr.find(f"{W}spacing")
        if spacing is not None:
            for attr in ("before", "after", "line", "lineRule", "beforeLines", "afterLines"):
                v = spacing.get(f"{W}{attr}")
                if v:
                    entry[f"spacing_{attr}"] = v
        ind = ppr.find(f"{W}ind")
        if ind is not None:
            for attr in ("left", "right", "firstLine", "hanging", "leftChars", "rightChars"):
                v = ind.get(f"{W}{attr}")
                if v:
                    entry[f"indent_{attr}"] = v
        shd = ppr.find(f"{W}shd")
        if shd is not None:
            if shd.get(f"{W}fill"):
                entry["shading_fill"] = shd.get(f"{W}fill")
            if shd.get(f"{W}val"):
                entry["shading_val"] = shd.get(f"{W}val")
        keep_next = ppr.find(f"{W}keepNext")
        if keep_next is not None:
            entry["keep_next"] = True
        page_break = ppr.find(f"{W}pageBreakBefore")
        if page_break is not None:
            entry["page_break_before"] = True
        outline = ppr.find(f"{W}outlineLvl")
        if outline is not None and outline.get(f"{W}val"):
            entry["outline_level"] = int(outline.get(f"{W}val"))
        # Borders
        pbdr = ppr.find(f"{W}pBdr")
        if pbdr is not None:
            borders = {}
            for side in ("top", "bottom", "left", "right", "between", "bar"):
                b = self._border_info(pbdr, f"{W}{side}")
                if b:
                    borders[side] = b
            if borders:
                entry["borders"] = borders
        return entry

    def extract_style_map(self) -> dict:
        """Comprehensive template style extraction.

        Captures everything: page layout, all paragraph styles (font, size,
        bold, italic, color, alignment, spacing, indent, borders), all table
        styles (borders, cell margins, column widths, alignment), and
        character styles. Called automatically on open_document.
        """
        style_map: dict = {
            "page": {},
            "default": {},
            "para_styles": {},
            "char_styles": {},
            "table_styles": [],
            "tables": [],
            "first_table_tblPr": None,
        }

        # ── styles.xml ────────────────────────────────────────────────────────
        styles = self._tree("word/styles.xml")
        if styles is not None:
            # Default (document-level) rPr/pPr
            doc_defaults = styles.find(f"{W}docDefaults")
            if doc_defaults is not None:
                rpr_def = doc_defaults.find(f".//{W}rPr")
                if rpr_def is not None:
                    style_map["default"].update(self._parse_rpr(rpr_def))
                ppr_def = doc_defaults.find(f".//{W}pPr")
                if ppr_def is not None:
                    style_map["default"].update(self._parse_ppr(ppr_def))

            for style in styles.iter(f"{W}style"):
                stype = style.get(f"{W}type", "")
                sid = style.get(f"{W}styleId", "")
                name_el = style.find(f"{W}name")
                name = name_el.get(f"{W}val", sid) if name_el is not None else sid

                entry: dict = {"styleId": sid}
                base = style.find(f"{W}basedOn")
                if base is not None and base.get(f"{W}val"):
                    entry["basedOn"] = base.get(f"{W}val")
                next_style = style.find(f"{W}next")
                if next_style is not None and next_style.get(f"{W}val"):
                    entry["next"] = next_style.get(f"{W}val")

                rpr = style.find(f"{W}rPr")
                if rpr is not None:
                    entry.update(self._parse_rpr(rpr))
                ppr = style.find(f"{W}pPr")
                if ppr is not None:
                    entry.update(self._parse_ppr(ppr))
                    # Also check rPr inside pPr (pPrChange etc.)
                    inner_rpr = ppr.find(f"{W}rPr")
                    if inner_rpr is not None:
                        entry.update(self._parse_rpr(inner_rpr))

                if stype == "paragraph":
                    style_map["para_styles"][name] = entry
                elif stype == "character":
                    style_map["char_styles"][name] = entry

            # Default = Normal style
            normal = style_map["para_styles"].get("Normal", {})
            if normal:
                style_map["default"].update(
                    {k: v for k, v in normal.items() if k not in style_map["default"]}
                )

        # ── document.xml ──────────────────────────────────────────────────────
        doc = self._tree("word/document.xml")
        if doc is not None:
            body = doc.find(f"{W}body")

            # ── Page layout ───────────────────────────────────────────────────
            if body is not None:
                sect_pr = body.find(f"{W}sectPr")
                if sect_pr is not None:
                    pg_sz = sect_pr.find(f"{W}pgSz")
                    if pg_sz is not None:
                        style_map["page"]["width_dxa"] = pg_sz.get(f"{W}w", "")
                        style_map["page"]["height_dxa"] = pg_sz.get(f"{W}h", "")
                        style_map["page"]["orientation"] = pg_sz.get(f"{W}orient", "portrait")
                    pg_mar = sect_pr.find(f"{W}pgMar")
                    if pg_mar is not None:
                        style_map["page"]["margins"] = {
                            k: pg_mar.get(f"{W}{k}", "")
                            for k in ("top", "bottom", "left", "right", "header", "footer", "gutter")
                            if pg_mar.get(f"{W}{k}")
                        }
                    cols = sect_pr.find(f"{W}cols")
                    if cols is not None:
                        style_map["page"]["columns"] = cols.get(f"{W}num", "1")
                        style_map["page"]["col_space"] = cols.get(f"{W}space", "")

            # ── Tables ────────────────────────────────────────────────────────
            for t_idx, tbl in enumerate(doc.iter(f"{W}tbl")):
                tinfo: dict = {"index": t_idx}
                tbl_pr = tbl.find(f"{W}tblPr")
                if tbl_pr is not None:
                    if style_map["first_table_tblPr"] is None:
                        style_map["first_table_tblPr"] = etree.tostring(tbl_pr)

                    ts = tbl_pr.find(f"{W}tblStyle")
                    if ts is not None:
                        tinfo["style"] = ts.get(f"{W}val", "")
                        if tinfo["style"] not in style_map["table_styles"]:
                            style_map["table_styles"].append(tinfo["style"])

                    jc = tbl_pr.find(f"{W}jc")
                    if jc is not None:
                        tinfo["alignment"] = jc.get(f"{W}val", "")

                    tw = tbl_pr.find(f"{W}tblW")
                    if tw is not None:
                        tinfo["width"] = {
                            "value": tw.get(f"{W}w", ""),
                            "type": tw.get(f"{W}type", ""),
                        }

                    # Borders
                    tbl_bdr = tbl_pr.find(f"{W}tblBorders")
                    if tbl_bdr is not None:
                        borders = {}
                        for side in ("top", "bottom", "left", "right", "insideH", "insideV"):
                            b = self._border_info(tbl_bdr, f"{W}{side}")
                            if b:
                                borders[side] = b
                        if borders:
                            tinfo["borders"] = borders

                    # Cell margins
                    tcmar = tbl_pr.find(f"{W}tblCellMar")
                    if tcmar is not None:
                        margins = {}
                        for side in ("top", "bottom", "left", "right"):
                            m = tcmar.find(f"{W}{side}")
                            if m is not None:
                                margins[side] = {
                                    "w": m.get(f"{W}w", ""),
                                    "type": m.get(f"{W}type", ""),
                                }
                        if margins:
                            tinfo["cell_margins"] = margins

                    # Table look flags
                    tbl_look = tbl_pr.find(f"{W}tblLook")
                    if tbl_look is not None:
                        tinfo["tbl_look"] = {
                            k: tbl_look.get(f"{W}{k}", "")
                            for k in ("val", "firstRow", "lastRow", "firstColumn", "lastColumn",
                                      "noHBand", "noVBand")
                            if tbl_look.get(f"{W}{k}")
                        }

                # Column widths from tblGrid
                tbl_grid = tbl.find(f"{W}tblGrid")
                if tbl_grid is not None:
                    tinfo["col_widths_dxa"] = [
                        gc.get(f"{W}w", "") for gc in tbl_grid.findall(f"{W}gridCol")
                    ]

                # First row cell properties (header row styling)
                rows = tbl.findall(f"{W}tr")
                if rows:
                    first_row_cells = []
                    for tc in rows[0].findall(f"{W}tc"):
                        tc_pr = tc.find(f"{W}tcPr")
                        cell_info: dict = {}
                        if tc_pr is not None:
                            tc_w = tc_pr.find(f"{W}tcW")
                            if tc_w is not None:
                                cell_info["width"] = {
                                    "value": tc_w.get(f"{W}w", ""),
                                    "type": tc_w.get(f"{W}type", ""),
                                }
                            tc_bdr = tc_pr.find(f"{W}tcBorders")
                            if tc_bdr is not None:
                                cborders = {}
                                for side in ("top", "bottom", "left", "right",
                                             "insideH", "insideV", "tl2br", "tr2bl"):
                                    b = self._border_info(tc_bdr, f"{W}{side}")
                                    if b:
                                        cborders[side] = b
                                if cborders:
                                    cell_info["borders"] = cborders
                            shd = tc_pr.find(f"{W}shd")
                            if shd is not None and shd.get(f"{W}fill"):
                                cell_info["shading"] = shd.get(f"{W}fill")
                            v_align = tc_pr.find(f"{W}vAlign")
                            if v_align is not None:
                                cell_info["v_align"] = v_align.get(f"{W}val", "")
                        first_row_cells.append(cell_info)
                    tinfo["first_row_cells"] = first_row_cells
                    tinfo["row_count"] = len(rows)

                style_map["tables"].append(tinfo)

        return style_map

    def apply_table_cell_style(self, table_idx: int, style_id: str) -> dict:
        """Apply a paragraph style to every cell in a table.

        Sets w:pStyle on each cell's paragraph so cells inherit correct
        font, size, spacing, and indentation from the named style.

        Args:
            table_idx: Zero-based table index.
            style_id: Style ID to apply (e.g. "TableText", "Table-Text").
        """
        doc = self._require("word/document.xml")
        tables = list(doc.iter(f"{W}tbl"))
        if table_idx < 0 or table_idx >= len(tables):
            raise IndexError(f"Table index {table_idx} out of range (have {len(tables)})")

        tbl = tables[table_idx]
        cells_updated = 0

        for tr in tbl.findall(f"{W}tr"):
            for tc in tr.findall(f"{W}tc"):
                for para in tc.findall(f"{W}p"):
                    ppr = para.find(f"{W}pPr")
                    if ppr is None:
                        ppr = etree.Element(f"{W}pPr")
                        para.insert(0, ppr)
                    # Remove existing pStyle
                    existing = ppr.find(f"{W}pStyle")
                    if existing is not None:
                        ppr.remove(existing)
                    # Insert new pStyle at position 0
                    ps = etree.Element(f"{W}pStyle")
                    ps.set(f"{W}val", style_id)
                    ppr.insert(0, ps)
                    cells_updated += 1

        self._mark("word/document.xml")
        return {"table_index": table_idx, "style_id": style_id, "cells_updated": cells_updated}

    def copy_table_style(self, source_idx: int, target_idx: int) -> dict:
        """Copy tblPr (borders, width, style) from one table to another.

        Args:
            source_idx: Table index to copy style FROM.
            target_idx: Table index to apply style TO.
        """
        doc = self._require("word/document.xml")
        tables = list(doc.iter(f"{W}tbl"))
        if source_idx < 0 or source_idx >= len(tables):
            raise IndexError(f"Source table {source_idx} out of range")
        if target_idx < 0 or target_idx >= len(tables):
            raise IndexError(f"Target table {target_idx} out of range")

        src_tbl = tables[source_idx]
        tgt_tbl = tables[target_idx]

        src_pr = src_tbl.find(f"{W}tblPr")
        if src_pr is None:
            raise ValueError(f"Source table {source_idx} has no tblPr")

        # Remove existing tblPr on target
        existing = tgt_tbl.find(f"{W}tblPr")
        if existing is not None:
            tgt_tbl.remove(existing)

        # Deep-copy source tblPr and insert at position 0
        new_pr = copy.deepcopy(src_pr)

        # Preserve jc (alignment) from target if it had one
        if existing is not None:
            jc = existing.find(f"{W}jc")
            if jc is not None:
                existing_jc = new_pr.find(f"{W}jc")
                if existing_jc is not None:
                    new_pr.remove(existing_jc)
                new_pr.append(copy.deepcopy(jc))

        tgt_tbl.insert(0, new_pr)
        self._mark("word/document.xml")
        return {"source": source_idx, "target": target_idx, "style_copied": True}

    def extract_template_structure(self) -> dict:
        """Extract full document structure: paragraphs and tables with indices.

        Returns:
            - paragraphs: list of {index, style, text (truncated 200 chars), paraId}
            - tables: list of {index, row_count, col_count, cells}

        Use this to map the template before filling.
        """
        doc = self._require("word/document.xml")
        body = doc.find(f"{W}body")

        paragraphs = []
        for idx, para in enumerate(body.iter(f"{W}p")):
            ppr = para.find(f"{W}pPr")
            style = ""
            if ppr is not None:
                ps = ppr.find(f"{W}pStyle")
                if ps is not None:
                    style = ps.get(f"{W}val", "")
            paragraphs.append(
                {
                    "index": idx,
                    "style": style,
                    "text": self._text(para)[:200],
                    "paraId": para.get(f"{W14}paraId", ""),
                }
            )

        tables = []
        for t_idx, tbl in enumerate(body.iter(f"{W}tbl")):
            rows = []
            for row in tbl.findall(f"{W}tr"):
                cells = [self._text(cell)[:100] for cell in row.findall(f"{W}tc")]
                rows.append(cells)
            tables.append(
                {
                    "index": t_idx,
                    "row_count": len(rows),
                    "col_count": len(rows[0]) if rows else 0,
                    "cells": rows,
                }
            )

        return {"paragraphs": paragraphs, "tables": tables}
