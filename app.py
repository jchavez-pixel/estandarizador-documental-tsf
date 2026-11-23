from __future__ import annotations

import json
import os
import hashlib
import re
import shutil
import subprocess
import tempfile
import zipfile
import unicodedata
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

import fitz
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import HTMLResponse, FileResponse
from starlette.background import BackgroundTask
from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph as RLParagraph, Spacer, Table as RLTable,
    TableStyle, Image as RLImage
)

APP_DIR = Path(__file__).resolve().parent
ASSET_LOGO = APP_DIR / "assets" / "tsf_logo.png"
GOLD = "F5BD00"
PALE_GOLD = "FFF4CC"
DARK = "242424"
MID = "626A73"
BORDER = "D7DCE2"
FONT = "Arial"
MAX_UPLOAD_MB = int(os.getenv("TSF_MAX_UPLOAD_MB", "50"))

app = FastAPI(title="Estandarizador Documental TSF", version="1.1.2")


def lima_today() -> str:
    try:
        return datetime.now(ZoneInfo("America/Lima")).strftime("%d/%m/%Y")
    except Exception:
        return datetime.now().strftime("%d/%m/%Y")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def docx_unsupported_features(path: Path) -> list[str]:
    """Flag complex Word structures that require a visual/manual review."""
    import zipfile
    flags = []
    try:
        with zipfile.ZipFile(path) as archive:
            xml = archive.read("word/document.xml").decode("utf-8", "ignore")
            checks = {
                "cuadros de texto flotantes": "w:txbxContent",
                "objetos OLE incrustados": "w:object",
                "contenido externo altChunk": "w:altChunk",
                "controles de contenido": "w:sdt",
                "marcas de revisión (insertado/eliminado)": "w:ins",
                "contenido eliminado por revisión": "w:del",
            }
            for label, token in checks.items():
                if token in xml:
                    flags.append(label)
            if any(name.startswith("word/embeddings/") for name in archive.namelist()):
                flags.append("archivos incrustados")
    except Exception:
        flags.append("estructura DOCX no inspeccionable")
    return flags


def clean_filename(value: str) -> str:
    value = Path(value).stem
    value = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode("ascii")
    value = re.sub(r"[^A-Za-z0-9._-]+", "_", value).strip("._-")
    return value[:90] or "documento"


def xml_set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)
    shd.set(qn("w:val"), "clear")


def xml_set_cell_border(cell, color: str = BORDER, size: str = "6") -> None:
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = "w:" + edge
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), size)
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), color)


def set_cell_text(cell, text: str, bold: bool = False, color: str = DARK, size: float = 8.5,
                  align=WD_ALIGN_PARAGRAPH.LEFT) -> None:
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    r = p.add_run(str(text or "Por confirmar"))
    r.bold = bold
    r.font.name = FONT
    r.font.size = Pt(size)
    r.font.color.rgb = RGBColor.from_string(color)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    xml_set_cell_border(cell)


def set_repeat_table_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def add_page_field(paragraph, field: str) -> None:
    run = paragraph.add_run()
    fld_char1 = OxmlElement("w:fldChar")
    fld_char1.set(qn("w:fldCharType"), "begin")
    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = f" {field} "
    fld_char2 = OxmlElement("w:fldChar")
    fld_char2.set(qn("w:fldCharType"), "end")
    run._r.append(fld_char1)
    run._r.append(instr_text)
    run._r.append(fld_char2)


def set_run_font(run, size=10.0, color=DARK, bold=None, italic=None) -> None:
    run.font.name = FONT
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), FONT)
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor.from_string(color)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic


def container_has_content(container) -> bool:
    for paragraph in container.paragraphs:
        if paragraph.text.strip():
            return True
        xml = paragraph._p.xml
        if any(token in xml for token in ("w:drawing", "w:object", "w:fldChar", "w:instrText")):
            return True
    return bool(container.tables)



def _clear_container(container) -> None:
    """Remove existing header/footer content so the corporate master is applied consistently."""
    element = container._element
    for child in list(element):
        element.remove(child)


def configure_section(section, title: str, code: str, version: str, status: str, logo_path: Path) -> dict[str, bool]:
    """Apply the TSF header/footer to every section, replacing legacy presentation elements."""
    section.page_width = Inches(8.2677)
    section.page_height = Inches(11.6929)
    section.top_margin = Inches(0.92)
    section.bottom_margin = Inches(0.78)
    section.left_margin = Inches(0.68)
    section.right_margin = Inches(0.68)
    section.header_distance = Inches(0.22)
    section.footer_distance = Inches(0.28)
    section.different_first_page_header_footer = False

    header = section.header
    footer = section.footer
    header_was_replaced = container_has_content(header)
    footer_was_replaced = container_has_content(footer)

    # Legacy headers/footers are presentation, not body provisions. Replace them with the TSF master.
    section.header.is_linked_to_previous = False
    section.footer.is_linked_to_previous = False
    _clear_container(header)
    _clear_container(footer)

    table = header.add_table(rows=1, cols=3, width=Inches(6.90))
    table.autofit = False
    widths = [Inches(1.42), Inches(3.68), Inches(1.80)]
    for idx, cell in enumerate(table.rows[0].cells):
        cell.width = widths[idx]
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        xml_set_cell_border(cell, color=GOLD, size="8")

    left = table.cell(0, 0)
    left.text = ""
    lp = left.paragraphs[0]
    lp.paragraph_format.space_after = Pt(0)
    lp.paragraph_format.space_before = Pt(0)
    if logo_path.exists():
        lp.add_run().add_picture(str(logo_path), width=Inches(1.25))
    else:
        run = lp.add_run("TOTAL\nSERVICIOS FINANCIEROS")
        set_run_font(run, 9, GOLD, True)

    middle = table.cell(0, 1)
    middle.text = ""
    mp = middle.paragraphs[0]
    mp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    mp.paragraph_format.space_after = Pt(0)
    run = mp.add_run(title.upper())
    set_run_font(run, 8.0, DARK, True)

    right = table.cell(0, 2)
    right.text = ""
    rp = right.paragraphs[0]
    rp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    rp.paragraph_format.space_after = Pt(0)
    run = rp.add_run(f"Código: {code or 'Por confirmar'}\n")
    set_run_font(run, 7.5, DARK, True)
    run = rp.add_run(f"Versión: {version or 'Por confirmar'}")
    set_run_font(run, 7.5, MID)

    # Word requires a paragraph after a table in a header for reliable editing/rendering.
    hp = header.add_paragraph()
    hp.paragraph_format.space_after = Pt(0)
    hp.paragraph_format.space_before = Pt(0)
    hp.paragraph_format.line_spacing = 0.1
    run = hp.add_run(" ")
    set_run_font(run, 1, "FFFFFF")

    fp = footer.add_paragraph()
    fp.paragraph_format.space_before = Pt(0)
    fp.paragraph_format.space_after = Pt(2)
    fp.paragraph_format.line_spacing = 1.0
    ppr = fp._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    top = OxmlElement("w:top")
    top.set(qn("w:val"), "single")
    top.set(qn("w:sz"), "5")
    top.set(qn("w:space"), "4")
    top.set(qn("w:color"), GOLD)
    p_bdr.append(top)
    ppr.append(p_bdr)
    run = fp.add_run(
        "Este documento es de uso interno y exclusivo de Total Servicios Financieros Empresa de Créditos. "
        "Su reproducción o distribución requiere autorización del área responsable."
    )
    set_run_font(run, 6.8, MID)

    p2 = footer.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after = Pt(0)
    run = p2.add_run(
        f"DOCUMENTO INTERNO · USO INTERNO Y CONFIDENCIAL | {status or 'PENDIENTE DE VALIDACIÓN'} | Página "
    )
    set_run_font(run, 7.0, DARK, True)
    add_page_field(p2, "PAGE")
    run = p2.add_run(" de ")
    set_run_font(run, 7.0, DARK, True)
    add_page_field(p2, "NUMPAGES")

    return {
        "header_preserved": False,
        "footer_preserved": False,
        "legacy_header_replaced": header_was_replaced,
        "legacy_footer_replaced": footer_was_replaced,
    }


def _iter_table_paragraphs(table):
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                yield paragraph
            for nested in cell.tables:
                yield from _iter_table_paragraphs(nested)


def _iter_all_paragraphs(doc):
    for paragraph in doc.paragraphs:
        yield paragraph, False
    for table in doc.tables:
        for paragraph in _iter_table_paragraphs(table):
            yield paragraph, True


def _iter_all_tables(doc):
    def walk(table):
        yield table
        for row in table.rows:
            for cell in row.cells:
                for nested in cell.tables:
                    yield from walk(nested)
    for table in doc.tables:
        yield from walk(table)


def _paragraph_is_heading(paragraph, text: str) -> tuple[bool, int]:
    style_name = paragraph.style.name.lower() if paragraph.style else ""
    if style_name.startswith("heading 1") or style_name in {"title", "subtitle"}:
        return True, 1
    if style_name.startswith("heading 2"):
        return True, 2
    if style_name.startswith("heading 3"):
        return True, 3

    compact = re.sub(r"\s+", " ", text).strip()
    is_caps = (
        len(compact) <= 125
        and compact.upper() == compact
        and any(char.isalpha() for char in compact)
    )
    numbered = bool(re.match(r"^(?:\d+(?:\.\d+){0,3}|[IVXLC]+)[.)]?\s+\S", compact, re.I)) and len(compact) <= 150
    direct_bold = bool(paragraph.runs) and all(
        (run.bold is True) for run in paragraph.runs if run.text.strip()
    )
    if is_caps or numbered:
        return True, 1 if is_caps or re.match(r"^\d+\s", compact) else 2
    if direct_bold and len(compact) <= 95:
        return True, 2
    return False, 0


def style_existing_content(doc: Document) -> None:
    """Normalize every body paragraph and every table, including nested tables."""
    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(9.6)
    normal.font.color.rgb = RGBColor.from_string(DARK)
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    normal.paragraph_format.space_after = Pt(4.5)
    normal.paragraph_format.line_spacing = 1.08

    for style_name, size, before, after in [
        ("Heading 1", 13, 11, 5),
        ("Heading 2", 11.5, 8, 4),
        ("Heading 3", 10.5, 6, 3),
    ]:
        try:
            style = styles[style_name]
            style.font.name = FONT
            style.font.size = Pt(size)
            style.font.bold = True
            style.font.color.rgb = RGBColor.from_string(DARK)
            style._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
            style.paragraph_format.space_before = Pt(before)
            style.paragraph_format.space_after = Pt(after)
            style.paragraph_format.keep_with_next = True
        except KeyError:
            pass

    for paragraph, in_table in _iter_all_paragraphs(doc):
        text = paragraph.text.strip()
        if not text:
            continue
        is_heading, level = _paragraph_is_heading(paragraph, text)
        for run in paragraph.runs:
            old_bold, old_italic, old_underline = run.bold, run.italic, run.underline
            if is_heading:
                size = {1: 11.2, 2: 10.5, 3: 9.8}.get(level, 10.5)
                set_run_font(run, size, DARK, True, old_italic)
            else:
                size = 8.6 if in_table else 9.6
                set_run_font(run, size, DARK, old_bold, old_italic)
                run.underline = old_underline

        if is_heading:
            paragraph.paragraph_format.keep_with_next = True
            paragraph.paragraph_format.space_before = Pt(7 if level == 1 else 5)
            paragraph.paragraph_format.space_after = Pt(3.5)
        else:
            if paragraph.paragraph_format.space_after is None:
                paragraph.paragraph_format.space_after = Pt(3 if in_table else 4.5)
            if not in_table:
                paragraph.paragraph_format.line_spacing = 1.08

    header_tokens = {
        "actividad", "etapa", "responsable", "campo", "n.º", "nº", "orden",
        "concepto", "versión", "version", "situación", "situacion", "disparador",
        "fecha", "contenido", "elaborador", "revisor", "aprobador", "criterio",
        "plazo", "requisito", "documento", "área", "area", "número", "numero",
        "tipo", "condición", "condicion", "validación", "validacion", "resultado",
        "rol", "cargo", "indicador", "observación", "observacion", "anexo", "check",
        "descripción", "descripcion", "paso", "tarea", "evidencia", "estado",
    }

    for table_index, table in enumerate(_iter_all_tables(doc)):
        table.autofit = True
        rows = table.rows
        if not rows:
            continue
        first_row_texts = [re.sub(r"\s+", " ", cell.text.strip().lower()) for cell in rows[0].cells]
        header_hits = sum(
            1 for value in first_row_texts
            if value and any(token in value for token in header_tokens)
        )
        is_header_row = len(first_row_texts) > 1 and (
            header_hits >= 2 or
            (first_row_texts[0] in header_tokens and header_hits >= 1)
        )
        metadata_pair_table = (
            table_index == 0 and len(first_row_texts) == 4 and
            first_row_texts[0] in {"código", "codigo", "campo"}
        )

        for row_index, row in enumerate(rows):
            if row_index == 0 and is_header_row:
                set_repeat_table_header(row)
            for col_index, cell in enumerate(row.cells):
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                xml_set_cell_border(cell, BORDER, "5")
                if row_index == 0 and is_header_row:
                    xml_set_cell_shading(cell, GOLD)
                elif metadata_pair_table and col_index in (0, 2):
                    xml_set_cell_shading(cell, PALE_GOLD)
                else:
                    # Remove legacy blue/gray fills from body cells without touching text.
                    tc_pr = cell._tc.get_or_add_tcPr()
                    shading = tc_pr.find(qn("w:shd"))
                    if shading is not None:
                        shading.set(qn("w:fill"), "FFFFFF")
                        shading.set(qn("w:val"), "clear")
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.space_before = Pt(0)
                    paragraph.paragraph_format.space_after = Pt(1.5)
                    paragraph.paragraph_format.line_spacing = 1.0
                    for run in paragraph.runs:
                        is_label = metadata_pair_table and col_index in (0, 2)
                        set_run_font(
                            run, 8.3, DARK,
                            True if (row_index == 0 and is_header_row) or is_label else run.bold,
                            run.italic,
                        )
                for nested in cell.tables:
                    # Nested tables are styled by _iter_all_tables above.
                    pass


def has_existing_document_cover(doc: Document) -> bool:
    first_text = " ".join(p.text.strip().upper() for p in doc.paragraphs[:10])
    if "INFORMACIÓN DEL DOCUMENTO" in first_text:
        return True
    if doc.tables:
        first_table = doc.tables[0]
        if len(first_table.rows) <= 12 and len(first_table.columns) >= 2:
            table_text = " ".join(cell.text.lower() for row in first_table.rows for cell in row.cells)
            has_code = "código" in table_text or "codigo" in table_text
            has_version = "versión" in table_text or "version" in table_text
            if has_code and has_version:
                return True
    return False


def existing_cover_metadata(doc: Document) -> dict[str, str]:
    """Read code/version from an existing metadata table without editing its text."""
    found: dict[str, str] = {}
    for table in doc.tables[:3]:
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            for index in range(0, len(cells) - 1, 2):
                label = re.sub(r"\s+", " ", cells[index].lower()).strip(" :")
                value = cells[index + 1].strip()
                if label in {"código", "codigo"} and value:
                    found.setdefault("code", value)
                elif label in {"versión", "version"} and value:
                    found.setdefault("version", value)
    return found


def add_cover_to_docx(doc: Document, meta: dict[str, str], logo_path: Path, cover_logo: bool = False) -> bool:
    # Do not duplicate an existing cover/control table. Its original content is preserved.
    if has_existing_document_cover(doc):
        return False

    body = doc._element.body
    original_first = next((child for child in list(body) if child.tag != qn("w:sectPr")), None)
    created = []

    if cover_logo and logo_path.exists():
        logo_p = doc.add_paragraph()
        logo_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        logo_p.paragraph_format.space_after = Pt(18)
        logo_p.add_run().add_picture(str(logo_path), width=Inches(1.55))
        created.append(logo_p._p)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(meta["doc_type"].upper())
    set_run_font(r, 9, MID, True)
    created.append(p._p)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(12)
    r = p.add_run(meta["title"])
    set_run_font(r, 21, DARK, True)
    created.append(p._p)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(12)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ppr = p._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "26")
    bottom.set(qn("w:space"), "2")
    bottom.set(qn("w:color"), GOLD)
    p_bdr.append(bottom)
    ppr.append(p_bdr)
    created.append(p._p)

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run("INFORMACIÓN DEL DOCUMENTO")
    set_run_font(r, 11, DARK, True)
    created.append(p._p)

    info_rows = [
        ("Código", meta["code"], "Documento", meta["title"]),
        ("Versión", meta["version"], "Fecha de estandarización", meta["issue_date"]),
        ("Área custodio", meta["owner_area"], "Macroproceso", meta["macroprocess"]),
        ("Proceso", meta["process"], "Subproceso", meta["subprocess"]),
        ("Estado", meta["status"], "Alcance", meta["scope"]),
        ("Antecedente", meta["antecedent"], "Tipo documental", meta["doc_type"]),
    ]
    table = doc.add_table(rows=len(info_rows), cols=4)
    table.autofit = False
    widths = [Inches(1.1), Inches(2.05), Inches(1.25), Inches(2.6)]
    for row_idx, row_data in enumerate(info_rows):
        for col_idx, val in enumerate(row_data):
            cell = table.cell(row_idx, col_idx)
            cell.width = widths[col_idx]
            is_label = col_idx in (0, 2)
            if is_label:
                xml_set_cell_shading(cell, PALE_GOLD)
            set_cell_text(cell, val or "Por confirmar", is_label, DARK, 8.0)
    created.append(table._tbl)

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(5)
    r = p.add_run("NIVELES DE APROBACIÓN")
    set_run_font(r, 11, DARK, True)
    created.append(p._p)

    approval = doc.add_table(rows=2, cols=3)
    approval.autofit = False
    for j, label in enumerate(["Elaborador", "Revisor", "Aprobador"]):
        approval.cell(0, j).width = Inches(2.33)
        xml_set_cell_shading(approval.cell(0, j), GOLD)
        set_cell_text(approval.cell(0, j), label, True, DARK, 8.5, WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_text(approval.cell(1, j), meta[["author", "reviewer", "approver"][j]] or "Por confirmar", False, MID, 8.0, WD_ALIGN_PARAGRAPH.CENTER)
    created.append(approval._tbl)

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(5)
    r = p.add_run("HISTORIAL DE CAMBIOS")
    set_run_font(r, 11, DARK, True)
    created.append(p._p)

    history = doc.add_table(rows=2, cols=4)
    history.autofit = True
    for j, label in enumerate(["Versión", "Fecha", "Aprobación", "Naturaleza del cambio"]):
        xml_set_cell_shading(history.cell(0, j), GOLD)
        set_cell_text(history.cell(0, j), label, True, DARK, 8.0)
    values = [meta["version"], meta["issue_date"], "Pendiente de aprobación", "Estandarización de formato; contenido fuente preservado."]
    for j, val in enumerate(values):
        set_cell_text(history.cell(1, j), val, False, DARK, 7.8)
    created.append(history._tbl)

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("CONTROL DOCUMENTAL")
    set_run_font(r, 9.5, DARK, True)
    created.append(p._p)

    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.12)
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run("Se estandariza la presentación del documento. El contenido original se conserva íntegramente y no se modifican sus disposiciones, cifras, responsables ni obligaciones. La revisión y aprobación corresponden al área responsable.")
    set_run_font(r, 8.2, MID)
    created.append(p._p)

    p = doc.add_paragraph()
    p.add_run().add_break(WD_BREAK.PAGE)
    created.append(p._p)

    if original_first is not None:
        for element in created:
            original_first.addprevious(element)
    else:
        # A blank Document() may contain only sectPr. Keep sectPr as the final body child.
        sect_pr = body.find(qn("w:sectPr"))
        if sect_pr is not None:
            for element in created:
                sect_pr.addprevious(element)
        else:
            for element in created:
                body.append(element)
    return True


def extract_docx_text(path: Path) -> str:
    doc = Document(path)
    pieces: list[str] = []
    for child in doc._element.body.iterchildren():
        if child.tag == qn("w:p"):
            pieces.append(Paragraph(child, doc).text)
        elif child.tag == qn("w:tbl"):
            table = Table(child, doc)
            for row in table.rows:
                pieces.append("\t".join(cell.text for cell in row.cells))
    return "\n".join(pieces)


def count_docx_images(doc: Document) -> int:
    return sum(1 for rel in doc.part.rels.values() if "image" in rel.reltype)


def process_docx(input_path: Path, output_docx: Path, meta: dict[str, str], report: dict[str, Any]) -> None:
    source_text = extract_docx_text(input_path)
    doc = Document(input_path)
    original_tables = len(doc.tables)
    original_paragraphs = len(doc.paragraphs)
    original_images = count_docx_images(doc)

    style_existing_content(doc)
    cover_added = add_cover_to_docx(doc, meta, ASSET_LOGO)
    existing_meta = existing_cover_metadata(doc) if not cover_added else {}
    header_meta = meta.copy()
    if existing_meta:
        # Do not rewrite source metadata. Reuse existing code/version in the repeated header.
        header_meta.update(existing_meta)
    section_preservation = []
    for section in doc.sections:
        section_preservation.append(configure_section(section, header_meta["title"], header_meta["code"], header_meta["version"], header_meta["status"], ASSET_LOGO))

    doc.core_properties.title = meta["title"]
    doc.core_properties.subject = "Documento estandarizado — contenido fuente preservado"
    doc.core_properties.author = "Estandarizador Documental TSF"
    doc.core_properties.keywords = "TSF, estandarización documental, control documental"
    doc.save(output_docx)

    output_text = extract_docx_text(output_docx)
    text_preserved = source_text in output_text if source_text else True
    reopened = Document(output_docx)
    output_tables = len(reopened.tables)
    output_images = count_docx_images(reopened)
    tables_preserved = output_tables >= original_tables
    images_preserved = output_images >= original_images
    source_features = docx_unsupported_features(input_path)
    output_features = docx_unsupported_features(output_docx)
    complex_features_preserved = all(item in output_features for item in source_features)
    integrity_pass = text_preserved and tables_preserved and images_preserved and complex_features_preserved
    report.update({
        "source_type": "DOCX",
        "cover_added": cover_added,
        "existing_cover_preserved": not cover_added,
        "existing_metadata_used_for_header": existing_meta,
        "original_headers_footers_preserved": section_preservation,
        "standard_header_footer_applied": True,
        "legacy_headers_replaced": any(x.get("legacy_header_replaced", False) for x in section_preservation),
        "legacy_footers_replaced": any(x.get("legacy_footer_replaced", False) for x in section_preservation),
        "source_paragraphs": original_paragraphs,
        "source_tables": original_tables,
        "source_images": original_images,
        "text_integrity_check": "PASS" if text_preserved else "REVIEW REQUIRED",
        "source_text_preserved": text_preserved,
        "tables_preserved": tables_preserved,
        "images_preserved": images_preserved,
        "complex_features_preserved": complex_features_preserved,
        "output_tables": output_tables,
        "output_images": output_images,
        "unsupported_features_requiring_manual_review": source_features,
        "manual_review_required": bool(source_features) or not cover_added or any(x["header_preserved"] or x["footer_preserved"] for x in section_preservation),
        "warnings": (["Se detectó una portada/tabla de control existente. Se conservó sin modificar; revise que el encabezado coincida con los datos de esa portada."] if not cover_added else []) + (["Se conservaron los encabezados/pies originales para no eliminar contenido fuente. Verifique su consistencia con la versión y los metadatos del formulario."] if any(x["header_preserved"] or x["footer_preserved"] for x in section_preservation) else []) + (["El archivo contiene estructuras complejas que requieren inspección visual manual: " + ", ".join(source_features)] if source_features else []),
        "integrity_status": "PASS" if integrity_pass else "REVIEW REQUIRED",
        "pdf_to_word_editability": "Not applicable — DOCX source retained as editable Word content.",
    })
    if not integrity_pass:
        raise RuntimeError("La comprobación de integridad del documento no fue satisfactoria. Se canceló la entrega automática.")



def find_word_executable() -> str | None:
    """Find Microsoft Word for Windows COM-based PDF export."""
    env_path = os.getenv("TSF_WORD_PATH")
    if env_path and Path(env_path).exists():
        return env_path
    if os.name != "nt":
        return None
    program_files = os.environ.get("ProgramFiles", r"C:\Program Files")
    program_files_x86 = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")
    candidates = [
        Path(program_files) / "Microsoft Office" / "root" / "Office16" / "WINWORD.EXE",
        Path(program_files_x86) / "Microsoft Office" / "root" / "Office16" / "WINWORD.EXE",
        Path(program_files) / "Microsoft Office" / "Office16" / "WINWORD.EXE",
        Path(program_files_x86) / "Microsoft Office" / "Office16" / "WINWORD.EXE",
        Path(r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"),
    ]
    for candidate in candidates:
        if candidate.exists():
            return str(candidate)
    return shutil.which("WINWORD.EXE") or shutil.which("winword")


def find_libreoffice() -> str | None:
    env_path = os.getenv("LIBREOFFICE_PATH")
    if env_path and Path(env_path).exists():
        return env_path
    for candidate in ("libreoffice", "soffice", "soffice.exe", "libreoffice.exe"):
        found = shutil.which(candidate)
        if found:
            return found
    windows_candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for candidate in windows_candidates:
        if Path(candidate).exists():
            return candidate
    return None


def find_pdf_export_engine() -> str | None:
    """Return the preferred available PDF export engine without starting it."""
    if os.name == "nt" and find_word_executable() and (shutil.which("powershell.exe") or shutil.which("powershell") or (Path(os.environ.get("WINDIR", r"C:\Windows")) / "System32" / "WindowsPowerShell" / "v1.0" / "powershell.exe").exists()):
        return "Microsoft Word COM"
    if find_libreoffice():
        return "LibreOffice"
    return None


def _word_com_export_pdf(docx_path: Path, pdf_path: Path, word_exe: str) -> None:
    """Use the user's installed Microsoft Word through PowerShell COM automation.

    This avoids pywin32 and does not require installing an additional Python package.
    Word is opened invisibly, the source DOCX is read-only, and ExportAsFixedFormat
    creates the PDF in the requested output directory.
    """
    powershell = (
        shutil.which("powershell.exe")
        or shutil.which("powershell")
        or str(Path(os.environ.get("WINDIR", r"C:\Windows")) / "System32" / "WindowsPowerShell" / "v1.0" / "powershell.exe")
    )
    if not Path(powershell).exists():
        raise RuntimeError("No se encontró Windows PowerShell para automatizar Microsoft Word.")

    script = r"""
$ErrorActionPreference = 'Stop'
$word = $null
$document = $null
$hadError = $false
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.ScreenUpdating = $false
    $document = $word.Documents.Open($env:TSF_DOCX_PATH, $false, $true, $false)
    $document.ExportAsFixedFormat($env:TSF_PDF_PATH, 17)
    $document.Close($false)
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($document)
    $document = $null
    $word.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    $word = $null
    Write-Output 'WORD_PDF_EXPORT_OK'
}
catch {
    [Console]::Error.WriteLine($_.Exception.ToString())
    $hadError = $true
}
finally {
    if ($null -ne $document) {
        try { $document.Close($false) } catch {}
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) } catch {}
    }
    if ($null -ne $word) {
        try { $word.Quit() } catch {}
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
if ($hadError) { exit 1 }
exit 0
"""
    import base64
    encoded_script = base64.b64encode(script.encode("utf-16le")).decode("ascii")
    env = os.environ.copy()
    env["TSF_DOCX_PATH"] = str(docx_path.resolve())
    env["TSF_PDF_PATH"] = str(pdf_path.resolve())
    env["TSF_WORD_EXE"] = word_exe
    if pdf_path.exists():
        pdf_path.unlink()
    try:
        completed = subprocess.run(
            [powershell, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-EncodedCommand", encoded_script],
            capture_output=True,
            text=True,
            timeout=240,
            env=env,
        )
    except subprocess.TimeoutExpired as exc:
        raise RuntimeError("Microsoft Word excedió el tiempo máximo de 240 segundos al generar el PDF.") from exc

    if completed.returncode != 0 or not pdf_path.exists() or pdf_path.stat().st_size == 0:
        detail = (completed.stderr or completed.stdout or "Sin detalle devuelto por PowerShell.").strip()
        raise RuntimeError("Microsoft Word no pudo generar el PDF mediante COM. " + detail[-1800:])


def _libreoffice_export_pdf(docx_path: Path, pdf_path: Path, executable: str) -> None:
    profile = pdf_path.parent / "lo_profile"
    profile.mkdir(exist_ok=True)
    cmd = [executable, "--headless", f"-env:UserInstallation={profile.as_uri()}", "--convert-to", "pdf", "--outdir", str(pdf_path.parent), str(docx_path)]
    completed = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
    if completed.returncode != 0 or not pdf_path.exists() or pdf_path.stat().st_size == 0:
        raise RuntimeError("LibreOffice no pudo generar el PDF. " + (completed.stderr[-1200:] or completed.stdout[-1200:]))


def docx_to_pdf(docx_path: Path, out_dir: Path, report: dict[str, Any] | None = None) -> Path:
    """Export DOCX to PDF using installed Word COM on Windows, or LibreOffice otherwise."""
    pdf_path = out_dir / (docx_path.stem + ".pdf")
    word_exe = find_word_executable() if os.name == "nt" else None
    if word_exe:
        try:
            _word_com_export_pdf(docx_path, pdf_path, word_exe)
            if report is not None:
                report["pdf_export"] = "Microsoft Word COM (ExportAsFixedFormat)"
            return pdf_path
        except Exception as word_error:
            executable = find_libreoffice()
            if not executable:
                raise RuntimeError(
                    "Se encontró Microsoft Word, pero falló la automatización COM al exportar el PDF. "
                    "No se requiere instalar LibreOffice. Detalle: " + str(word_error)
                ) from word_error
            try:
                _libreoffice_export_pdf(docx_path, pdf_path, executable)
                if report is not None:
                    report["pdf_export"] = "LibreOffice headless (respaldo tras fallo de Word COM)"
                return pdf_path
            except Exception as lo_error:
                raise RuntimeError(f"Fallaron Microsoft Word COM y LibreOffice. Word: {word_error}. LibreOffice: {lo_error}") from lo_error

    executable = find_libreoffice()
    if executable:
        _libreoffice_export_pdf(docx_path, pdf_path, executable)
        if report is not None:
            report["pdf_export"] = "LibreOffice headless"
        return pdf_path

    raise RuntimeError(
        "No se encontró un motor de exportación PDF disponible. En Windows se buscó Microsoft Word COM; "
        "también puede utilizarse LibreOffice si ya está instalado. Verifique que Microsoft Word esté instalado "
        "y que la automatización COM esté habilitada para su usuario."
    )


def rl_paragraph(text: str, style: ParagraphStyle) -> RLParagraph:
    safe = (text or "Por confirmar").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return RLParagraph(safe, style)


def build_pdf_cover(path: Path, meta: dict[str, str]) -> None:
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("TSFTitle", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=20, leading=25, textColor=colors.HexColor("#242424"), alignment=TA_CENTER, spaceAfter=10)
    label_style = ParagraphStyle("TSFLabel", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=8.3, leading=10, textColor=colors.HexColor("#242424"))
    body_style = ParagraphStyle("TSFBody", parent=styles["Normal"], fontName="Helvetica", fontSize=8.3, leading=10, textColor=colors.HexColor("#242424"))
    small_style = ParagraphStyle("TSFSmall", parent=styles["Normal"], fontName="Helvetica", fontSize=7.6, leading=9, textColor=colors.HexColor("#626A73"))
    heading_style = ParagraphStyle("TSFHeading", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=10, leading=12, textColor=colors.HexColor("#242424"), spaceBefore=8, spaceAfter=5)
    doc = SimpleDocTemplate(str(path), pagesize=A4, rightMargin=17*mm, leftMargin=17*mm, topMargin=16*mm, bottomMargin=17*mm)
    story = []
    if ASSET_LOGO.exists():
        logo = RLImage(str(ASSET_LOGO), width=45*mm, height=10*mm)
        story.append(logo)
    story.append(Spacer(1, 16*mm))
    story.append(RLParagraph(meta["doc_type"].upper(), ParagraphStyle("Type", parent=small_style, alignment=TA_CENTER, fontName="Helvetica-Bold", fontSize=9)))
    story.append(Spacer(1, 3*mm))
    story.append(rl_paragraph(meta["title"], title_style))
    story.append(Spacer(1, 4*mm))
    info_title = rl_paragraph("INFORMACIÓN DEL DOCUMENTO", heading_style)
    story.append(info_title)
    pairs = [
        ("Código", meta["code"], "Documento", meta["title"]),
        ("Versión", meta["version"], "Fecha de estandarización", meta["issue_date"]),
        ("Área custodio", meta["owner_area"], "Macroproceso", meta["macroprocess"]),
        ("Proceso", meta["process"], "Subproceso", meta["subprocess"]),
        ("Estado", meta["status"], "Alcance", meta["scope"]),
        ("Antecedente", meta["antecedent"], "Tipo documental", meta["doc_type"]),
    ]
    data = []
    for a, b, c, d in pairs:
        data.append([rl_paragraph(a, label_style), rl_paragraph(b, body_style), rl_paragraph(c, label_style), rl_paragraph(d, body_style)])
    info = RLTable(data, colWidths=[27*mm, 53*mm, 34*mm, 62*mm], repeatRows=0)
    info.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (0,-1), colors.HexColor("#FFF4CC")),
        ("BACKGROUND", (2,0), (2,-1), colors.HexColor("#FFF4CC")),
        ("GRID", (0,0), (-1,-1), 0.45, colors.HexColor("#D7DCE2")),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("LEFTPADDING", (0,0), (-1,-1), 6), ("RIGHTPADDING", (0,0), (-1,-1), 6),
        ("TOPPADDING", (0,0), (-1,-1), 6), ("BOTTOMPADDING", (0,0), (-1,-1), 6),
    ]))
    story.append(info)
    story.append(Spacer(1, 6*mm))
    story.append(rl_paragraph("NIVELES DE APROBACIÓN", heading_style))
    approval_data = [[rl_paragraph(x, label_style) for x in ["Elaborador", "Revisor", "Aprobador"]],
                     [rl_paragraph(meta["author"], body_style), rl_paragraph(meta["reviewer"], body_style), rl_paragraph(meta["approver"], body_style)]]
    approval = RLTable(approval_data, colWidths=[58*mm, 58*mm, 60*mm])
    approval.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#F5BD00")),
        ("GRID", (0,0), (-1,-1), 0.45, colors.HexColor("#D7DCE2")),
        ("ALIGN", (0,0), (-1,-1), "CENTER"), ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("TOPPADDING", (0,0), (-1,-1), 7), ("BOTTOMPADDING", (0,0), (-1,-1), 7),
    ]))
    story.append(approval)
    story.append(Spacer(1, 6*mm))
    story.append(rl_paragraph("HISTORIAL DE CAMBIOS", heading_style))
    history_data = [
        [rl_paragraph(x, label_style) for x in ["Versión", "Fecha", "Aprobación", "Naturaleza del cambio"]],
        [rl_paragraph(meta["version"], body_style), rl_paragraph(meta["issue_date"], body_style), rl_paragraph("Pendiente de aprobación", body_style), rl_paragraph("Estandarización de formato; contenido fuente preservado.", body_style)],
    ]
    history = RLTable(history_data, colWidths=[24*mm, 30*mm, 42*mm, 80*mm])
    history.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#F5BD00")),
        ("GRID", (0,0), (-1,-1), 0.45, colors.HexColor("#D7DCE2")),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("TOPPADDING", (0,0), (-1,-1), 6), ("BOTTOMPADDING", (0,0), (-1,-1), 6),
    ]))
    story.append(history)
    story.append(Spacer(1, 5*mm))
    story.append(rl_paragraph("CONTROL DOCUMENTAL", heading_style))
    story.append(rl_paragraph("Se estandariza la presentación del documento. El contenido fuente se conserva íntegramente; la revisión y aprobación corresponden al área responsable.", small_style))

    def draw_footer(canvas, document):
        canvas.saveState()
        w, h = A4
        canvas.setStrokeColor(colors.HexColor("#8A8A8A"))
        canvas.setLineWidth(0.5)
        canvas.line(17*mm, 13*mm, w-17*mm, 13*mm)
        canvas.setFont("Helvetica", 6.8)
        canvas.setFillColor(colors.HexColor("#626A73"))
        canvas.drawString(17*mm, 8.5*mm, "Este documento es de uso interno y exclusivo de Total Servicios Financieros Empresa de Créditos.")
        canvas.setFont("Helvetica-Bold", 7)
        canvas.setFillColor(colors.HexColor("#242424"))
        canvas.drawCentredString(w/2, 4.5*mm, f"DOCUMENTO INTERNO · USO INTERNO Y CONFIDENCIAL | {meta['status']} | Página {document.page}")
        canvas.restoreState()
    doc.build(story, onFirstPage=draw_footer)



def _pdf_word_counter(text: str) -> Counter:
    words = re.findall(r"[0-9A-Za-zÁÉÍÓÚÜÑáéíóúüñ]+", (text or "").casefold())
    return Counter(words)


def convert_pdf_to_editable_docx(input_pdf: Path, converted_docx: Path, report: dict[str, Any]) -> None:
    """Convert a text-based PDF into editable Word content before TSF styling."""
    pdf = fitz.open(input_pdf)
    if pdf.is_encrypted:
        raise RuntimeError("El PDF está protegido con contraseña. No se puede reconstruir de forma segura.")
    source_text = "\n".join(page.get_text("text", sort=True) for page in pdf)
    page_count = len(pdf)
    pdf.close()

    compact_text = re.sub(r"\s+", "", source_text)
    if len(compact_text) < 25 or len(_pdf_word_counter(source_text)) < 4:
        raise RuntimeError(
            "El PDF parece escaneado o no contiene texto seleccionable. "
            "No se generó una salida parcial. Se requiere incorporar OCR para este tipo de archivo."
        )

    try:
        from pdf2docx import Converter
    except ImportError as exc:
        raise RuntimeError(
            "Falta el motor de reconstrucción PDF→Word. En la terminal del proyecto ejecuta: "
            "python -m pip install --only-binary=:all: -r requirements.txt"
        ) from exc

    converter = Converter(str(input_pdf))
    try:
        converter.convert(str(converted_docx), start=0, end=None, multi_processing=False)
    finally:
        converter.close()

    if not converted_docx.exists() or converted_docx.stat().st_size == 0:
        raise RuntimeError("El motor PDF→Word no produjo un documento editable.")

    converted_text = extract_docx_text(converted_docx)
    source_words = _pdf_word_counter(source_text)
    converted_words = _pdf_word_counter(converted_text)
    total_source_words = sum(source_words.values())
    missing_words = sum((source_words - converted_words).values())
    token_recall = 1.0 if total_source_words == 0 else max(0.0, 1.0 - missing_words / total_source_words)

    report.update({
        "pdf_source_pages": page_count,
        "pdf_source_text_characters": len(source_text),
        "pdf_converted_text_characters": len(converted_text),
        "pdf_source_word_count": total_source_words,
        "pdf_conversion_word_recall": round(token_recall, 4),
        "pdf_conversion_engine": "pdf2docx",
    })

    if token_recall < 0.97:
        raise RuntimeError(
            "La conversión del PDF no superó el control de integridad del texto "
            f"(retención estimada: {token_recall:.1%}). No se entregó el archivo para evitar omisiones. "
            "Revisa si el PDF contiene tablas complejas, columnas o texto como imagen."
        )


def process_pdf(
    input_path: Path,
    output_docx: Path,
    output_pdf: Path,
    meta: dict[str, str],
    report: dict[str, Any],
    work_dir: Path,
) -> None:
    """Rebuild PDF content as editable DOCX, apply TSF styling, then export the new PDF."""
    converted_source = work_dir / "pdf_reconstruido_editable.docx"
    convert_pdf_to_editable_docx(input_path, converted_source, report)

    # Reuse the same full-document DOCX styling engine used for native Word files.
    process_docx(converted_source, output_docx, meta, report)

    generated_pdf = docx_to_pdf(output_docx, work_dir, report)
    if generated_pdf.resolve() != output_pdf.resolve():
        shutil.copy2(generated_pdf, output_pdf)

    report.update({
        "source_type": "PDF",
        "pdf_source_sha256": sha256_file(input_path),
        "pdf_output_method": "PDF→DOCX editable reconstruction→TSF styling→PDF",
        "pdf_source_pages_preserved_unchanged": False,
        "pdf_pages_reflowed_to_corporate_layout": True,
        "word_output_editability": (
            "Editable Word reconstruction. Text and detected tables are editable; "
            "images and complex vector elements may require visual review."
        ),
        "manual_review_required": True,
        "warnings": list(report.get("warnings", [])) + [
            "El PDF se reconstruyó como contenido editable y se volvió a maquetar. "
            "Verificar visualmente tablas, imágenes, diagramas, columnas y anexos."
        ],
    })


HTML = r'''<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="theme-color" content="#171717">
<title>Estandarizador Documental TSF</title>
<style>
:root{
  --gold:#f5bd00;--gold-dark:#d9a600;--ink:#202124;--muted:#68717c;
  --line:#dfe3e8;--bg:#f3f5f8;--card:#fff;--soft:#fffaf0;
  --green:#147a4b;--red:#b42318;--shadow:0 8px 24px rgba(20,28,40,.055);
}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font-family:Arial,Helvetica,sans-serif}
.top{background:#151515;color:#fff;border-bottom:5px solid var(--gold)}
.top-inner{max-width:1060px;margin:auto;padding:18px 24px;display:flex;align-items:center;gap:20px}
.logo{width:172px;max-width:35vw;background:#fff;padding:5px 8px;border-radius:4px;object-fit:contain}
.brand-copy h1{font-size:20px;line-height:1.2;margin:0;font-weight:750;letter-spacing:.1px}
.brand-copy p{margin:6px 0 0;color:#d1d5db;font-size:12px}
.version{margin-left:auto;border:1px solid #555;border-radius:999px;padding:7px 11px;color:#f3f4f6;font-size:11px;white-space:nowrap}
main{max-width:960px;margin:34px auto 42px;padding:0 20px}
.hero{margin:0 0 22px}
.eyebrow{display:inline-flex;align-items:center;gap:8px;color:#745900;background:#fff4cc;border:1px solid #f2d777;border-radius:999px;padding:6px 11px;font-size:10px;font-weight:800;letter-spacing:1px;text-transform:uppercase}
.eyebrow:before{content:"";width:7px;height:7px;background:var(--gold);border-radius:50%}
.hero h2{font-size:27px;line-height:1.2;margin:13px 0 8px;letter-spacing:-.45px}
.hero p{font-size:14px;line-height:1.6;color:var(--muted);margin:0;max-width:760px}
.card{background:var(--card);border:1px solid var(--line);border-radius:15px;padding:23px 25px;margin:0 0 16px;box-shadow:var(--shadow)}
.section-head{display:flex;align-items:center;gap:11px;margin:0 0 17px}
.step{display:inline-flex;align-items:center;justify-content:center;width:29px;height:29px;flex:0 0 29px;background:#fff4cc;color:#654f00;border:1px solid #f1d56c;border-radius:9px;font-size:12px;font-weight:800}
.section-head h3{font-size:15px;margin:0;font-weight:750}
.section-head p{font-size:11px;color:var(--muted);margin:4px 0 0}
.dropzone{position:relative;display:flex;align-items:center;gap:17px;border:1.5px dashed #d7aa00;border-radius:12px;background:linear-gradient(135deg,#fffdf6,#fff9e8);padding:21px 22px;cursor:pointer;transition:border-color .18s,background .18s,transform .18s}
.dropzone:hover,.dropzone.dragover{border-color:#b88900;background:#fff5d6}
.dropzone.has-file{border-style:solid;border-color:#d5b13a;background:#fffaf0}
.upload-icon{width:48px;height:48px;display:flex;align-items:center;justify-content:center;flex:0 0 48px;background:#fff0b8;border-radius:13px;color:#5d4900}
.upload-copy{min-width:0;flex:1}
.upload-copy strong{display:block;font-size:14px;margin-bottom:5px}
.upload-copy span{display:block;color:var(--muted);font-size:12px;line-height:1.45}
.file-name{margin-top:5px!important;color:#6b5200!important;font-weight:700;overflow-wrap:anywhere}
.select-file{display:inline-flex;align-items:center;justify-content:center;background:var(--gold);color:#171717;border-radius:8px;padding:10px 14px;font-size:12px;font-weight:800;white-space:nowrap}
#file{position:absolute;width:1px;height:1px;opacity:0;pointer-events:none}
.grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:17px 20px}
.grid.responsibles{grid-template-columns:repeat(3,minmax(0,1fr))}
.field{display:flex;flex-direction:column;gap:7px;min-width:0}
.field label{font-size:12px;font-weight:750;color:#30343a}
.field label .required{color:#b42318;margin-left:3px}
.field label .optional{font-size:10px;font-weight:600;color:#7b8490;margin-left:5px}
.field input,.field select{width:100%;height:46px;border:1px solid #cfd5dd;border-radius:9px;padding:0 13px;font-size:13px;background:#fff;color:#202124;outline:none;transition:border-color .15s,box-shadow .15s}
.field input::placeholder{color:#9aa2ad}
.field input:focus,.field select:focus{border-color:#c99a00;box-shadow:0 0 0 3px rgba(245,189,0,.16)}
.field small{font-size:10px;color:#89919b;line-height:1.4}
.form-note{display:flex;align-items:flex-start;gap:10px;margin-top:17px;padding:12px 14px;border-radius:9px;background:#f7f8fa;color:#69717b;font-size:11px;line-height:1.5}
.form-note svg{flex:0 0 16px;margin-top:1px;color:#987400}
.actions{display:flex;align-items:center;justify-content:space-between;gap:18px;flex-wrap:wrap}
.action-copy strong{display:block;font-size:14px;margin-bottom:4px}
.action-copy span{display:block;font-size:11px;color:var(--muted);line-height:1.45}
.btn{display:inline-flex;align-items:center;justify-content:center;gap:9px;min-height:46px;border:0;border-radius:9px;background:var(--gold);color:#171717;font-size:13px;font-weight:800;padding:0 21px;cursor:pointer;box-shadow:0 3px 0 #d5a500;transition:transform .15s,background .15s}
.btn:hover{background:#ffca1c;transform:translateY(-1px)}
.btn:disabled{opacity:.65;cursor:wait;transform:none}
.status{min-height:18px;margin-top:13px;color:var(--muted);font-size:12px;line-height:1.5}
.status.success{color:var(--green);font-weight:700}
.status.error{color:var(--red);font-weight:700}
.footer{text-align:center;color:#8b929c;font-size:10px;padding:8px 16px 22px}
@media(max-width:760px){
  .top-inner{padding:15px 18px;gap:13px}.logo{width:128px}.brand-copy h1{font-size:17px}.version{display:none}
  main{margin:24px auto 30px;padding:0 14px}.hero h2{font-size:23px}
  .card{padding:19px 17px;border-radius:12px}.grid.responsibles{grid-template-columns:repeat(2,minmax(0,1fr))}
  .dropzone{padding:17px;gap:12px}.select-file{padding:9px 10px}
}
@media(max-width:520px){
  .top-inner{align-items:flex-start}.logo{width:105px}.brand-copy h1{font-size:15px}.brand-copy p{font-size:10px}
  .grid,.grid.responsibles{grid-template-columns:1fr;gap:14px}
  .dropzone{align-items:flex-start;flex-wrap:wrap}.upload-icon{width:40px;height:40px;flex-basis:40px}
  .select-file{margin-left:52px}.actions{align-items:stretch}.btn{width:100%}
}
</style>
</head>
<body>
<header class="top">
  <div class="top-inner">
    <img class="logo" src="/logo" alt="TOTAL Servicios Financieros">
    <div class="brand-copy">
      <h1>Estandarizador Documental TSF</h1>
      <p>Motor de estandarización corporativa · Procesamiento temporal</p>
    </div>
    <span class="version">VERSIÓN 1.1.2</span>
  </div>
</header>

<main>
  <div class="hero">
    <span class="eyebrow">Documentos corporativos</span>
    <h2>Estandariza tu documento</h2>
    <p>Carga el archivo original, completa los datos esenciales y genera una versión Word editable y un PDF con el formato corporativo TSF.</p>
  </div>

  <form id="form" enctype="multipart/form-data">
    <section class="card">
      <div class="section-head">
        <span class="step">01</span>
        <div><h3>Documento de origen</h3><p>Selecciona el archivo que deseas estandarizar.</p></div>
      </div>
      <label class="dropzone" id="dropzone" for="file">
        <input id="file" required type="file" name="file" accept=".docx,.pdf,application/pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document">
        <span class="upload-icon" aria-hidden="true">
          <svg width="25" height="25" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 16V4"/><path d="m7 9 5-5 5 5"/><path d="M4 16.5v2A1.5 1.5 0 0 0 5.5 20h13a1.5 1.5 0 0 0 1.5-1.5v-2"/></svg>
        </span>
        <span class="upload-copy">
          <strong>Arrastra tu archivo aquí o selecciónalo</strong>
          <span>Formatos admitidos: Word (.docx) y PDF · Tamaño máximo: 50 MB</span>
          <span class="file-name" id="file-name">Ningún archivo seleccionado</span>
        </span>
        <span class="select-file">Seleccionar archivo</span>
      </label>
    </section>

    <section class="card">
      <div class="section-head">
        <span class="step">02</span>
        <div><h3>Datos del documento</h3><p>Solo necesitamos la clasificación y el nombre oficial.</p></div>
      </div>
      <div class="grid">
        <div class="field">
          <label for="doc_type">Tipo documental<span class="required">*</span></label>
          <select id="doc_type" name="doc_type" required>
            <option value="" selected disabled>Selecciona el tipo documental</option>
            <option value="Política">Política</option>
            <option value="Procedimiento">Procedimiento</option>
            <option value="Manual">Manual</option>
            <option value="Instructivo">Instructivo</option>
            <option value="Reglamento">Reglamento</option>
            <option value="Otro">Otro</option>
          </select>
        </div>
        <div class="field">
          <label for="title">Nombre del documento<span class="required">*</span></label>
          <input id="title" required name="title" maxlength="180" placeholder="Escribe el nombre oficial del documento">
        </div>
      </div>
    </section>

    <section class="card">
      <div class="section-head">
        <span class="step">03</span>
        <div><h3>Responsables y aprobación</h3><p>Identifica a quienes elaboran y revisan el documento.</p></div>
      </div>
      <div class="grid responsibles">
        <div class="field">
          <label for="author">Elaborador<span class="required">*</span></label>
          <input id="author" required name="author" maxlength="140" placeholder="Nombre y cargo">
          <small>Campo obligatorio</small>
        </div>
        <div class="field">
          <label for="reviewer">Revisor<span class="required">*</span></label>
          <input id="reviewer" required name="reviewer" maxlength="140" placeholder="Nombre y cargo">
          <small>Campo obligatorio</small>
        </div>
        <div class="field">
          <label for="approver">Aprobador<span class="optional">OPCIONAL</span></label>
          <input id="approver" name="approver" maxlength="140" placeholder="Nombre y cargo">
          <small>Completa este campo solo cuando corresponda.</small>
        </div>
      </div>
      <div class="form-note">
        <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.8"><circle cx="12" cy="12" r="9"/><path d="M12 11v5m0-8h.01"/></svg>
        <span>Los campos marcados con <b>*</b> son obligatorios. El aprobador puede dejarse en blanco. El documento generado queda sujeto a revisión del área responsable.</span>
      </div>
    </section>

    <section class="card">
      <div class="actions">
        <div class="action-copy">
          <strong>Generar documentos estandarizados</strong>
          <span>Recibirás un paquete ZIP con Word editable, PDF y reporte de integridad.</span>
        </div>
        <button class="btn" id="submit" type="submit">
          <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 3v12m-5-5 5 5 5-5"/><path d="M5 17v3h14v-3"/></svg>
          Estandarizar documento
        </button>
      </div>
      <div id="status" class="status" role="status" aria-live="polite"></div>
    </section>
  </form>
</main>
<div class="footer">TOTAL Servicios Financieros · Estandarizador Documental TSF · v1.1.2</div>

<script>
const form = document.getElementById('form');
const btn = document.getElementById('submit');
const statusBox = document.getElementById('status');
const fileInput = document.getElementById('file');
const dropzone = document.getElementById('dropzone');
const fileName = document.getElementById('file-name');
const maxBytes = 50 * 1024 * 1024;

function showSelectedFile(file) {
  if (!file) {
    fileName.textContent = 'Ningún archivo seleccionado';
    dropzone.classList.remove('has-file');
    return;
  }
  fileName.textContent = `${file.name} · ${(file.size / 1024 / 1024).toFixed(2)} MB`;
  dropzone.classList.add('has-file');
}

fileInput.addEventListener('change', () => showSelectedFile(fileInput.files[0]));
['dragenter','dragover'].forEach(eventName => dropzone.addEventListener(eventName, event => {
  event.preventDefault();
  dropzone.classList.add('dragover');
}));
['dragleave','drop'].forEach(eventName => dropzone.addEventListener(eventName, event => {
  event.preventDefault();
  dropzone.classList.remove('dragover');
}));
dropzone.addEventListener('drop', event => {
  const files = event.dataTransfer.files;
  if (files.length) {
    const transfer = new DataTransfer();
    transfer.items.add(files[0]);
    fileInput.files = transfer.files;
    showSelectedFile(files[0]);
  }
});

form.addEventListener('submit', async event => {
  event.preventDefault();
  if (!form.reportValidity()) return;

  const selected = fileInput.files[0];
  if (!selected) return;
  if (selected.size > maxBytes) {
    statusBox.className = 'status error';
    statusBox.textContent = 'El archivo supera el límite de 50 MB.';
    return;
  }

  btn.disabled = true;
  btn.textContent = 'Procesando documento…';
  statusBox.className = 'status';
  statusBox.textContent = 'Analizando el archivo y aplicando el formato corporativo. No cierres esta ventana…';

  try {
    const response = await fetch('/process', {method:'POST', body:new FormData(form)});
    if (!response.ok) {
      let message = 'No se pudo procesar el archivo.';
      try {
        const result = await response.json();
        message = result.detail || message;
      } catch (_) {}
      throw new Error(message);
    }

    const blob = await response.blob();
    const objectUrl = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = objectUrl;
    link.download = response.headers.get('X-Output-Filename') || 'Estandarizador_TSF_resultados.zip';
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(objectUrl), 1500);

    statusBox.className = 'status success';
    statusBox.textContent = 'Proceso completado. Se descargó el paquete con Word, PDF y reporte de integridad.';
  } catch (error) {
    statusBox.className = 'status error';
    statusBox.textContent = error.message;
  } finally {
    btn.disabled = false;
    btn.innerHTML = '<svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 3v12m-5-5 5 5 5-5"/><path d="M5 17v3h14v-3"/></svg> Estandarizar documento';
  }
});
</script>
</body>
</html>'''


@app.get("/", response_class=HTMLResponse)
def home():
    return HTML.replace("__DATE__", lima_today())


@app.get("/logo")
def logo():
    if not ASSET_LOGO.exists():
        raise HTTPException(404, "Logo no encontrado")
    return FileResponse(ASSET_LOGO, media_type="image/png")


@app.get("/health")
def health():
    return {"status": "ok", "app": "Estandarizador Documental TSF", "version": app.version,
            "pdf_engine": find_pdf_export_engine(), "word_executable": bool(find_word_executable()),
            "libreoffice": bool(find_libreoffice())}


@app.post("/process")
async def process(
    file: UploadFile = File(...),
    doc_type: str = Form(...),
    title: str = Form(...),
    author: str = Form(...),
    reviewer: str = Form(...),
    approver: str = Form(""),
):
    doc_type = doc_type.strip()
    title = title.strip()
    author = author.strip()
    reviewer = reviewer.strip()
    approver = approver.strip()

    if not doc_type:
        raise HTTPException(400, "Selecciona el tipo documental.")
    if not title:
        raise HTTPException(400, "El nombre del documento es obligatorio.")
    if not author:
        raise HTTPException(400, "El elaborador es obligatorio.")
    if not reviewer:
        raise HTTPException(400, "El revisor es obligatorio.")

    filename = file.filename or "documento"
    suffix = Path(filename).suffix.lower()
    if suffix not in {".docx", ".pdf"}:
        raise HTTPException(400, "Formato no admitido. Carga un archivo .docx o .pdf.")
    if not title.strip():
        raise HTTPException(400, "El nombre del documento es obligatorio.")
    data = await file.read()
    if not data:
        raise HTTPException(400, "El archivo está vacío.")
    if len(data) > MAX_UPLOAD_MB * 1024 * 1024:
        raise HTTPException(413, f"El archivo supera el límite de {MAX_UPLOAD_MB} MB.")

    work_dir = Path(tempfile.mkdtemp(prefix="tsf_estandarizador_"))
    base = clean_filename(filename)
    source_path = work_dir / (base + suffix)
    source_path.write_bytes(data)
    output_docx = work_dir / f"{base}_TSF.docx"
    output_pdf = work_dir / f"{base}_TSF.pdf"
    report: dict[str, Any] = {
        "application": "Estandarizador Documental TSF",
        "application_version": app.version,
        "source_filename": filename,
        "processed_at": datetime.now().isoformat(timespec="seconds"),
        "policy": "Preservación del contenido fuente; estandarización integral de presentación en todas las páginas.",
        "metadata": {},
    }
    meta = {
        "doc_type": doc_type,
        "title": title,
        "code": "Por confirmar",
        "version": "Por confirmar",
        "issue_date": lima_today(),
        "antecedent": "No consignado",
        "owner_area": "Por confirmar",
        "macroprocess": "Por confirmar",
        "process": "Por confirmar",
        "subprocess": "Por confirmar",
        "status": "Pendiente de validación",
        "scope": "Por confirmar",
        "author": author,
        "reviewer": reviewer,
        "approver": approver or "No consignado",
    }
    report["metadata"] = meta
    report["source_sha256"] = sha256_file(source_path)
    try:
        if suffix == ".docx":
            process_docx(source_path, output_docx, meta, report)
            pdf_path = docx_to_pdf(output_docx, work_dir, report)
            if pdf_path != output_pdf:
                shutil.copy2(pdf_path, output_pdf)
        else:
            process_pdf(source_path, output_docx, output_pdf, meta, report, work_dir)

        report["output_docx_sha256"] = sha256_file(output_docx)
        report["output_pdf_sha256"] = sha256_file(output_pdf)
        report_path = work_dir / "reporte_integridad.json"
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        zip_path = work_dir / f"{base}_Estandarizado_TSF.zip"
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
            z.write(output_docx, output_docx.name)
            z.write(output_pdf, output_pdf.name)
            z.write(report_path, report_path.name)
        download_name = zip_path.name
        return FileResponse(
            zip_path,
            media_type="application/zip",
            filename=download_name,
            headers={"X-Output-Filename": download_name},
            background=BackgroundTask(shutil.rmtree, work_dir, ignore_errors=True),
        )
    except HTTPException:
        shutil.rmtree(work_dir, ignore_errors=True)
        raise
    except Exception as exc:
        shutil.rmtree(work_dir, ignore_errors=True)
        raise HTTPException(500, str(exc))


def main():
    import uvicorn
    host = os.getenv("TSF_HOST", "127.0.0.1")
    port = int(os.getenv("TSF_PORT", "8765"))
    print("=" * 62)
    print(" ESTANDARIZADOR DOCUMENTAL TSF · v1.1.2")
    print(f" Abrir en el navegador: http://{host}:{port}")
    print(" El servicio se ejecuta localmente; no publica en SharePoint.")
    print("=" * 62)
    uvicorn.run(app, host=host, port=port, log_level="info")


if __name__ == "__main__":
    main()
