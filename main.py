"""
Offerte docx-extractie service.

Doel: een geüpload .docx bestand omzetten naar leesbare ruwe tekst,
met correcte afhandeling van samengevoegde tabelcellen (het patroon
dat we bij de Brink-offerte tegenkwamen: een hele tabel die eigenlijk
één grote merged cell is met spatie-uitlijning).

Dit vervangt GEEN semantische parsing (welk getal hoort bij welk item) -
dat blijft de taak van de Claude-call verderop in de n8n flow. Deze
service zorgt alleen dat er schone, correcte tekst uit het bestand komt
om aan die Claude-call te voeren, zonder dat python-docx eigenaardigheden
(gedupliceerde merged cells, kapotte tabellen) de output vervuilen.

Draai lokaal:
    pip install -r requirements.txt
    uvicorn main:app --host 0.0.0.0 --port 8000

Endpoint:
    POST /extract-docx
    multipart/form-data, veld "file" = het .docx bestand

    Response:
    {
      "raw_text": "... alle uitgelezen tekst ...",
      "tables_found": 2,
      "tables_used": 1,
      "warnings": ["Tabel 2 was volledig leeg, overgeslagen."]
    }
"""

import io
from fastapi import FastAPI, UploadFile, File, HTTPException
from docx import Document

app = FastAPI(title="Offerte docx-extractie service")


def dedupe_row_cells(row):
    """Geef elke onderliggende tabelcel van een rij precies één keer terug,
    ook als de rij horizontaal samengevoegde cellen bevat (die anders
    meerdere keren achter elkaar worden geretourneerd door python-docx)."""
    seen = set()
    cells = []
    for cell in row.cells:
        cid = id(cell._tc)
        if cid in seen:
            continue
        seen.add(cid)
        cells.append(cell)
    return cells


def extract_table_text(table):
    """Zet een hele tabel om naar leesbare tekst, rij voor rij, met de
    cellen van elke rij gescheiden door ' | '. Slaat volledig lege rijen
    over (zoals de 37 lege rijen die we bij Brink tegenkwamen)."""
    lines = []
    for row in table.rows:
        cells = dedupe_row_cells(row)
        cell_texts = [c.text.strip() for c in cells]
        if not any(cell_texts):
            continue
        lines.append(" | ".join(cell_texts))
    return "\n".join(lines)


@app.post("/extract-docx")
async def extract_docx(file: UploadFile = File(...)):
    if not file.filename.lower().endswith((".docx",)):
        raise HTTPException(status_code=400, detail="Alleen .docx bestanden worden ondersteund door dit endpoint.")

    content = await file.read()
    try:
        doc = Document(io.BytesIO(content))
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Kon het bestand niet openen als docx: {e}")

    warnings = []
    used_blocks = []

    # Tekst buiten tabellen (voorblad-teksten, losse notities e.d.)
    body_text = "\n".join(p.text.strip() for p in doc.paragraphs if p.text.strip())
    if body_text:
        used_blocks.append("TEKST BUITEN TABELLEN:\n" + body_text)

    tables_used = 0
    for i, table in enumerate(doc.tables):
        table_text = extract_table_text(table)
        if not table_text.strip():
            warnings.append(f"Tabel {i + 1} was volledig leeg, overgeslagen.")
            continue
        used_blocks.append(f"TABEL {i + 1}:\n" + table_text)
        tables_used += 1

    raw_text = "\n\n---\n\n".join(used_blocks)

    if not raw_text.strip():
        warnings.append("Geen bruikbare tekst gevonden in het hele document.")

    return {
        "raw_text": raw_text,
        "tables_found": len(doc.tables),
        "tables_used": tables_used,
        "warnings": warnings,
    }


@app.get("/health")
async def health():
    return {"status": "ok"}


# ---------------------------------------------------------------------------
# Generatie: JSON (output van de Claude-parse-stap) -> Word-document
# ---------------------------------------------------------------------------

import io as _io
import base64
from typing import Optional, List
from pydantic import BaseModel
from fastapi.responses import StreamingResponse
from docx import Document as DocxDocument
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FONT_NAME = "Calibri"
HEADER_BG = "2F5233"
SHADE_BG = "F2F2F2"
TOTAL_BG = "E4E4E4"


class Regel(BaseModel):
    item: str
    omschrijving: Optional[str] = ""
    specs: Optional[List[str]] = []
    aantal: Optional[float] = None
    prijs_per_stuk: Optional[float] = None
    totaal: float
    waarschuwing: Optional[str] = None
    opmerking: Optional[str] = None


class Klant(BaseModel):
    naam: Optional[str] = None
    adres: Optional[str] = None
    contact: Optional[str] = None
    logo_base64: Optional[str] = None  # ruwe base64, zonder data:-prefix


class GenerateRequest(BaseModel):
    regels: List[Regel]
    algemene_opmerkingen: Optional[List[str]] = []
    klant: Optional[Klant] = None
    template: Optional[str] = "1"


def _shade_cell(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_hex)
    tcPr.append(shd)


def _repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    header = OxmlElement("w:tblHeader")
    header.set(qn("w:val"), "true")
    trPr.append(header)


def _set_font(run, size=10, bold=False, color=None):
    run.font.name = FONT_NAME
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def _money(n):
    return f"€ {n:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def build_docx_template_1(data: GenerateRequest) -> bytes:
    doc = DocxDocument()
    section = doc.sections[0]
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)

    if data.klant and data.klant.logo_base64:
        try:
            logo_bytes = base64.b64decode(data.klant.logo_base64)
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run()
            run.add_picture(_io.BytesIO(logo_bytes), width=Cm(4))
        except Exception:
            pass  # kapotte/ontbrekende logo-data mag de generatie niet laten crashen

    if data.klant and data.klant.naam:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(f"Offerte voor {data.klant.naam}")
        _set_font(run, size=16, bold=True)
        if data.klant.adres:
            p2 = doc.add_paragraph()
            p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run2 = p2.add_run(data.klant.adres)
            _set_font(run2, size=10)
        if data.klant.contact:
            p3 = doc.add_paragraph()
            p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run3 = p3.add_run(data.klant.contact)
            _set_font(run3, size=10)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("Offerte")
    _set_font(run, size=18, bold=True)

    table = doc.add_table(rows=1, cols=5)
    table.autofit = False
    widths = [Cm(3.5), Cm(6), Cm(1.8), Cm(2.5), Cm(2.7)]
    headers = ["Item", "Omschrijving / specificaties", "Aantal", "Prijs p.s.", "Totaal"]

    hdr_cells = table.rows[0].cells
    for i, (htext, w) in enumerate(zip(headers, widths)):
        hdr_cells[i].width = w
        _shade_cell(hdr_cells[i], HEADER_BG)
        hdr_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = hdr_cells[i].paragraphs[0]
        run = p.add_run(htext)
        _set_font(run, size=10, bold=True, color="FFFFFF")
    _repeat_header(table.rows[0])

    grand_total = 0.0
    for idx, regel in enumerate(data.regels):
        row = table.add_row()
        shade = idx % 2 == 1
        for i, w in enumerate(widths):
            row.cells[i].width = w
            if shade:
                _shade_cell(row.cells[i], SHADE_BG)

        p = row.cells[0].paragraphs[0]
        run = p.add_run(regel.item)
        _set_font(run, bold=True)

        cell1 = row.cells[1]
        cell1.paragraphs[0].text = ""
        first = True
        if regel.omschrijving:
            par = cell1.paragraphs[0] if first else cell1.add_paragraph()
            run = par.add_run(regel.omschrijving)
            _set_font(run)
            first = False
        for spec in (regel.specs or []):
            par = cell1.paragraphs[0] if first else cell1.add_paragraph()
            run = par.add_run(f"•  {spec}")
            _set_font(run, size=9, color="444444")
            first = False
        if regel.opmerking:
            par = cell1.paragraphs[0] if first else cell1.add_paragraph()
            run = par.add_run(regel.opmerking)
            _set_font(run, size=9, color="555555")
            run.italic = True
            first = False
        if first:
            cell1.paragraphs[0].add_run("")

        p = row.cells[2].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run(str(int(regel.aantal)) if regel.aantal is not None else "")
        _set_font(run)

        p = row.cells[3].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run(_money(regel.prijs_per_stuk) if regel.prijs_per_stuk is not None else "")
        _set_font(run)

        p = row.cells[4].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run(_money(regel.totaal))
        _set_font(run, bold=True)

        grand_total += regel.totaal

    total_row = table.add_row()
    total_row.cells[0].merge(total_row.cells[3])
    _shade_cell(total_row.cells[0], TOTAL_BG)
    _shade_cell(total_row.cells[4], TOTAL_BG)
    p = total_row.cells[0].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run("TOTAAL")
    _set_font(run, size=11, bold=True)
    p = total_row.cells[4].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run(_money(grand_total))
    _set_font(run, size=11, bold=True)

    if data.algemene_opmerkingen:
        doc.add_paragraph()
        for note in data.algemene_opmerkingen:
            p = doc.add_paragraph()
            run = p.add_run(note)
            _set_font(run, size=9, color="555555")
            run.italic = True

    buf = _io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


TEMPLATE_BUILDERS = {
    "1": build_docx_template_1,
    "2": build_docx_template_1,  # TODO: eigen compacte variant
    "3": build_docx_template_1,  # TODO: eigen huisstijl-variant
}


@app.post("/generate-docx")
async def generate_docx(data: GenerateRequest):
    builder = TEMPLATE_BUILDERS.get(data.template or "1")
    if builder is None:
        raise HTTPException(status_code=400, detail=f"Onbekend template: {data.template}")

    docx_bytes = builder(data)
    filename = "offerte.docx"
    return StreamingResponse(
        _io.BytesIO(docx_bytes),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )
