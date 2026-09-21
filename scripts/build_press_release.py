#!/usr/bin/env python3
"""Build the 2026 winners press-release document.

Reads data/awards.yml and data/winners/*.yml and writes a Word document that
matches the layout and styling of the 2024 release kept in reference/.
Winners whose arbitration is not finished yet are rendered as placeholder rows,
so the document can be regenerated and handed over at any point.

    python3 scripts/build_press_release.py [-o output/path.docx]
"""

import argparse
import os
import pathlib
import sys

import yaml
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

ROOT = pathlib.Path(__file__).resolve().parent.parent
DATA = ROOT / "data"

# A winner is shown with real content once arbitration is final.
RENDERED_STATUSES = {"confirmed", "drafted", "approved"}


# --------------------------------------------------------------------------- #
# low-level OOXML helpers (python-docx has no direct API for RTL or shading)
# --------------------------------------------------------------------------- #
def _tag(name, **attrs):
    el = OxmlElement(name)
    for key, value in attrs.items():
        el.set(qn(f"w:{key}"), value)
    return el


def rtl_paragraph(par, align=None, space_after=0):
    """Mark a paragraph as right-to-left and tighten its spacing."""
    pPr = par._p.get_or_add_pPr()
    pPr.append(_tag("w:bidi"))
    par.paragraph_format.space_after = Pt(space_after)
    par.paragraph_format.space_before = Pt(0)
    if align is not None:
        par.alignment = align
    return par


def shade_paragraph(par, fill):
    par._p.get_or_add_pPr().append(
        _tag("w:shd", val="clear", color="auto", fill=fill)
    )


def clear_cell(cell):
    """Leave the cell with exactly one empty, run-free paragraph."""
    for par in cell.paragraphs[1:]:
        par._p.getparent().remove(par._p)
    par = cell.paragraphs[0]
    for run in par.runs:
        run._r.getparent().remove(run._r)
    return par


def shade_cell(cell, fill):
    cell._tc.get_or_add_tcPr().append(
        _tag("w:shd", val="clear", color="auto", fill=fill)
    )


def style_run(run, font, size_pt, bold=False, color=None, rtl=False):
    run.font.name = font
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    rPr = run._r.get_or_add_rPr()
    rPr.rFonts.set(qn("w:cs"), font)          # complex-script font, for Arabic
    szCs = _tag("w:szCs", val=str(int(size_pt * 2)))
    rPr.append(szCs)
    if bold:
        rPr.append(_tag("w:bCs"))
    if rtl:
        rPr.append(_tag("w:rtl"))
    return run


def table_borders(table, color):
    tblPr = table._tbl.tblPr
    tblPr.append(_tag("w:bidiVisual"))        # right-to-left column order
    borders = OxmlElement("w:tblBorders")
    for edge, size in (
        ("top", "12"), ("left", "12"), ("bottom", "12"), ("right", "12"),
        ("insideH", "6"), ("insideV", "6"),
    ):
        borders.append(_tag(f"w:{edge}", val="single", sz=size, space="0", color=color))
    tblPr.append(borders)


# --------------------------------------------------------------------------- #
# document building
# --------------------------------------------------------------------------- #
def add_text_cell(cell, text, style, rtl, align=None, valign=False):
    """Fill a table cell, keeping blank-line-separated paragraphs intact.

    Within a paragraph the source line wrapping is reflowed away, so YAML block
    scalars can be wrapped for readability without breaking lines in Word.
    """
    blocks = [" ".join(b.split()) for b in str(text).split("\n\n") if b.strip()] or [""]
    first = clear_cell(cell)
    for index, block in enumerate(blocks):
        par = first if index == 0 else cell.add_paragraph()
        if rtl:
            rtl_paragraph(par, align)
        else:
            par.paragraph_format.space_after = Pt(0)
            if align is not None:
                par.alignment = align
        style_run(par.add_run(block), style["font"], style["body_pt"], rtl=rtl)
    if valign:
        cell._tc.get_or_add_tcPr().append(_tag("w:vAlign", val="center"))


def add_award_table(doc, award, winners, cfg):
    style = cfg["style"]
    widths = [Inches(w) for w in (0.45, 1.9, 1.8, 5.2)]   # ratios from the 2024 file

    table = doc.add_table(rows=0, cols=4)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table_borders(table, style["border"])

    # row 1: award name, merged across all four columns, on the teal band
    row = table.add_row()
    title_cell = row.cells[0].merge(row.cells[3])
    par = rtl_paragraph(clear_cell(title_cell), WD_ALIGN_PARAGRAPH.CENTER)
    style_run(par.add_run(award["name_ar"]), style["font"], style["award_pt"],
              bold=True, color="FFFFFF", rtl=True)
    shade_cell(title_cell, style["teal"])

    # row 2: column headers, same teal band
    row = table.add_row()
    for cell, heading, width in zip(row.cells, cfg["columns_ar"], widths):
        cell.width = width
        par = rtl_paragraph(clear_cell(cell), WD_ALIGN_PARAGRAPH.CENTER)
        style_run(par.add_run(heading), style["font"], style["body_pt"],
                  bold=True, color="FFFFFF", rtl=True)
        shade_cell(cell, style["teal"])

    # data rows, one per winner slot
    for number, winner in enumerate(winners, start=1):
        ready = winner.get("status") in RENDERED_STATUSES
        name = winner.get("name_ar") or ""
        notes = winner.get("notes_ar") or ""
        bio = winner.get("bio_en") or ""
        if not ready or not name:
            name = cfg["placeholder_ar"]
            notes = notes or "—"
            bio = bio or "—"

        row = table.add_row()
        cells = row.cells
        for cell, width in zip(cells, widths):
            cell.width = width
        add_text_cell(cells[0], str(number), style, rtl=True,
                      align=WD_ALIGN_PARAGRAPH.CENTER, valign=True)
        add_text_cell(cells[1], name, style, rtl=True)
        add_text_cell(cells[2], notes, style, rtl=True)
        add_text_cell(cells[3], bio, style, rtl=False)

    return table


def build(cfg, winners_by_award, out_path):
    style = cfg["style"]
    doc = Document()

    normal = doc.styles["Normal"]
    normal.font.name = style["font"]
    normal.font.size = Pt(style["body_pt"])
    normal.element.rPr.rFonts.set(qn("w:cs"), style["font"])

    section = doc.sections[0]
    section.page_width, section.page_height = Inches(8.5), Inches(11)
    for side in ("top", "bottom", "left", "right"):
        setattr(section, f"{side}_margin", Inches(0.5))

    par = rtl_paragraph(doc.add_paragraph(), WD_ALIGN_PARAGRAPH.CENTER)
    style_run(par.add_run(cfg["title_ar"]), style["font"], style["title_pt"],
              bold=True, rtl=True)

    for sec in cfg["sections"]:
        rtl_paragraph(doc.add_paragraph())
        par = rtl_paragraph(doc.add_paragraph(), WD_ALIGN_PARAGRAPH.CENTER)
        shade_paragraph(par, style["gold"])
        style_run(par.add_run(sec["heading_ar"]), style["font"], style["section_pt"],
                  bold=True, color="FFFFFF", rtl=True)

        for award in sec["awards"]:
            rtl_paragraph(doc.add_paragraph())
            add_award_table(doc, award, winners_by_award[award["id"]], cfg)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out_path)


# --------------------------------------------------------------------------- #
def load_winners(cfg):
    """Return {award_id: [winner, ...]} ordered by slot, padded to winner_count."""
    loaded = {}
    for path in sorted((DATA / "winners").glob("*.yml")):
        winner = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
        winner["_path"] = path.name
        loaded.setdefault(winner.get("award"), []).append(winner)

    by_award, problems = {}, []
    for sec in cfg["sections"]:
        for award in sec["awards"]:
            slots = loaded.get(award["id"], [])
            if len(slots) != award["winner_count"]:
                problems.append(
                    f"{award['id']}: found {len(slots)} winner files, "
                    f"expected {award['winner_count']}"
                )
            by_award[award["id"]] = slots[: award["winner_count"]] + [
                {"status": "pending"} for _ in range(award["winner_count"] - len(slots))
            ]
    return by_award, problems


def report(cfg, by_award):
    print(f"\n  {cfg['title_ar']}\n")
    total = ready = 0
    for sec in cfg["sections"]:
        print(f"  {sec['heading_ar']}")
        for award in sec["awards"]:
            slots = by_award[award["id"]]
            done = sum(
                1 for w in slots
                if w.get("status") in RENDERED_STATUSES and w.get("name_ar")
            )
            written = sum(1 for w in slots if (w.get("bio_en") or "").strip())
            total += len(slots)
            ready += done
            print(f"    {award['name_en']:<45} names {done}/{len(slots)}   "
                  f"write-ups {written}/{len(slots)}")
        print()
    print(f"  {ready}/{total} winners confirmed\n")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("-o", "--output", type=pathlib.Path,
                        default=ROOT / "output" / "winners-press-release-2026.docx")
    args = parser.parse_args()

    cfg = yaml.safe_load((DATA / "awards.yml").read_text(encoding="utf-8"))
    by_award, problems = load_winners(cfg)
    for problem in problems:
        print(f"  warning: {problem}", file=sys.stderr)

    build(cfg, by_award, args.output)
    report(cfg, by_award)
    print(f"  written to {os.path.relpath(args.output, ROOT)}\n")


if __name__ == "__main__":
    main()
