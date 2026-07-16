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
import zipfile
import base64
from fastapi import FastAPI, UploadFile, File, HTTPException
from docx import Document

try:
    from PIL import Image
except ImportError:
    Image = None

app = FastAPI(title="Offerte docx-extractie service")

MIN_IMAGE_DIMENSION = 60      # kleinere afbeeldingen (bullets, iconen) worden overgeslagen
THUMBNAIL_MAX_WIDTH = 400     # basis64-thumbnails blijven klein, geen volledige resolutie nodig voor een keuzelijst
MAX_IMAGES_RETURNED = 8


def extract_embedded_images(file_bytes):
    """Haal alle ingebedde afbeeldingen uit een docx (via word/media/), zodat de
    gebruiker er zelf een logo uit kan kiezen in plaats van dat het systeem raadt
    welke afbeelding de logo is. Slaat te kleine afbeeldingen (iconen, bullets) over."""
    images = []
    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
            media_files = sorted(n for n in z.namelist() if n.startswith("word/media/"))
            for name in media_files:
                if len(images) >= MAX_IMAGES_RETURNED:
                    break
                raw = z.read(name)
                if Image is None:
                    continue
                try:
                    img = Image.open(io.BytesIO(raw))
                    w, h = img.size
                except Exception:
                    continue
                if w < MIN_IMAGE_DIMENSION or h < MIN_IMAGE_DIMENSION:
                    continue
                if w > THUMBNAIL_MAX_WIDTH:
                    ratio = THUMBNAIL_MAX_WIDTH / w
                    img = img.convert("RGB").resize((THUMBNAIL_MAX_WIDTH, int(h * ratio)))
                buf = io.BytesIO()
                img.convert("RGB").save(buf, format="JPEG", quality=80)
                images.append({
                    "filename": name.rsplit("/", 1)[-1],
                    "width": w,
                    "height": h,
                    "thumbnail_base64": base64.b64encode(buf.getvalue()).decode(),
                })
    except zipfile.BadZipFile:
        pass
    return images


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

    images = extract_embedded_images(content)
    if not images:
        warnings.append("Geen (bruikbare) afbeeldingen gevonden om als logo te kiezen.")

    return {
        "raw_text": raw_text,
        "tables_found": len(doc.tables),
        "tables_used": tables_used,
        "images": images,
        "warnings": warnings,
    }


@app.get("/health")
async def health():
    return {"status": "ok"}


# ---------------------------------------------------------------------------
# Generatie: JSON (output van de Claude-parse-stap) -> Word-document
# ---------------------------------------------------------------------------

import io as _io
import os
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
REFURNITY_LOGO_PATH = os.path.join(os.path.dirname(__file__), "refurnity_logo.png")
REFURNITY_ADDRESS_LINES = [
    "Refurnity BV",
    "Bergerweg 6, 6085 AT Horn",
    "Telefoon 0621502536",
    "Mail info@ReFurnity.nl",
    "www.ReFurnity.nl",
]


class Regel(BaseModel):
    item: str
    omschrijving: Optional[str] = ""
    specs: Optional[List[str]] = []
    aantal: Optional[float] = None
    prijs_per_stuk: Optional[float] = None
    totaal: float
    waarschuwing: Optional[str] = None
    opmerking: Optional[str] = None
    foto_base64: Optional[str] = None  # optionele productfoto, ruwe base64 zonder data:-prefix


class Klant(BaseModel):
    naam: Optional[str] = None
    adres: Optional[str] = None
    contact: Optional[str] = None
    logo_base64: Optional[str] = None  # ruwe base64, zonder data:-prefix


class AanbetalingTermijn(BaseModel):
    label: str
    percentage: float  # bijv. 50 voor 50%


class GenerateRequest(BaseModel):
    regels: List[Regel]
    algemene_opmerkingen: Optional[List[str]] = []
    klant: Optional[Klant] = None
    template: Optional[str] = "1"
    toeslag_percentage: Optional[float] = None  # bijv. 20 voor 20% opslag op het leveranciersbedrag
    btw_percentage: Optional[float] = 21  # standaard 21%, zet expliciet op 0 om BTW-regel weg te laten
    aanbetaling_termijnen: Optional[List[AanbetalingTermijn]] = None


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

    # Voorblad krijgt geen kop-logo (daar staat al een groot logo in de tekst zelf).
    # Vanaf pagina 2 verschijnt automatisch het kleine logo rechtsboven.
    section.different_first_page_header_footer = True
    if os.path.exists(REFURNITY_LOGO_PATH):
        header = section.header
        hp = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hrun = hp.add_run()
        hrun.add_picture(REFURNITY_LOGO_PATH, width=Cm(2.8))
        # lege first-page header, zodat het voorblad zelf geen kop-logo krijgt
        section.first_page_header.paragraphs[0].text = ""

    # --- Voorblad ---
    if data.klant and data.klant.logo_base64:
        try:
            logo_bytes = base64.b64decode(data.klant.logo_base64)
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run()
            run.add_picture(_io.BytesIO(logo_bytes), width=Cm(6))
        except Exception:
            pass  # kapotte/ontbrekende logo-data mag de generatie niet laten crashen

    if data.klant and data.klant.naam:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(f"Offerte voor {data.klant.naam}")
        _set_font(run, size=18, bold=True)
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

    # ReFurnity-logo en -adres, groot en gecentreerd, alleen op het voorblad
    doc.add_paragraph()
    if os.path.exists(REFURNITY_LOGO_PATH):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(REFURNITY_LOGO_PATH, width=Cm(7))
    for line in REFURNITY_ADDRESS_LINES:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(line)
        _set_font(run, size=9, color="666666")

    doc.add_page_break()

    # --- Inhoudspagina ---
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

        if regel.foto_base64:
            try:
                foto_bytes = base64.b64decode(regel.foto_base64)
                fp = row.cells[0].add_paragraph()
                frun = fp.add_run()
                frun.add_picture(_io.BytesIO(foto_bytes), width=Cm(2.8))
            except Exception:
                pass  # kapotte fotodata mag de generatie niet laten crashen

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

    def _add_summary_row(label, bedrag, bold=True, shade_bg=None):
        row = table.add_row()
        row.cells[0].merge(row.cells[3])
        if shade_bg:
            _shade_cell(row.cells[0], shade_bg)
            _shade_cell(row.cells[4], shade_bg)
        p = row.cells[0].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run(label)
        _set_font(run, size=11 if bold else 10, bold=bold)
        p = row.cells[4].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run(_money(bedrag))
        _set_font(run, size=11 if bold else 10, bold=bold)

    subtotaal = grand_total
    heeft_toeslag = bool(data.toeslag_percentage)
    heeft_btw = bool(data.btw_percentage)

    if not heeft_toeslag and not heeft_btw:
        _add_summary_row("TOTAAL", subtotaal, shade_bg=TOTAL_BG)
        eindtotaal = subtotaal
    else:
        if heeft_toeslag:
            _add_summary_row("Subtotaal", subtotaal, bold=False)
            toeslag_bedrag = subtotaal * (data.toeslag_percentage / 100)
            _add_summary_row(f"Toeslag ({data.toeslag_percentage:g}%)", toeslag_bedrag, bold=False)
            totaal_excl_btw = subtotaal + toeslag_bedrag
        else:
            totaal_excl_btw = subtotaal

        if heeft_btw:
            _add_summary_row("Totaal excl. BTW", totaal_excl_btw, bold=False)
            btw_bedrag = totaal_excl_btw * (data.btw_percentage / 100)
            _add_summary_row(f"BTW ({data.btw_percentage:g}%)", btw_bedrag, bold=False)
            eindtotaal = totaal_excl_btw + btw_bedrag
            _add_summary_row("Totaal incl. BTW", eindtotaal, shade_bg=TOTAL_BG)
        else:
            eindtotaal = totaal_excl_btw
            _add_summary_row("TOTAAL", eindtotaal, shade_bg=TOTAL_BG)

    if data.aanbetaling_termijnen:
        doc.add_paragraph()
        p = doc.add_paragraph()
        run = p.add_run("Betaaltermijnen")
        _set_font(run, size=11, bold=True)
        for termijn in data.aanbetaling_termijnen:
            bedrag = eindtotaal * (termijn.percentage / 100)
            p = doc.add_paragraph()
            run = p.add_run(f"{termijn.label}: {termijn.percentage:g}% = {_money(bedrag)}")
            _set_font(run, size=10)

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


import subprocess
import tempfile
import uuid


def convert_docx_to_pdf(docx_bytes: bytes) -> bytes:
    """Converteer docx-bytes naar pdf-bytes via headless LibreOffice.
    Elke aanroep gebruikt een eigen tijdelijke map, zodat gelijktijdige
    requests elkaar niet in de weg zitten."""
    with tempfile.TemporaryDirectory() as tmpdir:
        docx_path = os.path.join(tmpdir, f"{uuid.uuid4().hex}.docx")
        with open(docx_path, "wb") as f:
            f.write(docx_bytes)

        result = subprocess.run(
            [
                "soffice", "--headless", "--norestore",
                "--convert-to", "pdf", "--outdir", tmpdir, docx_path,
            ],
            capture_output=True, text=True, timeout=60,
        )
        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice-conversie mislukt: {result.stderr}")

        pdf_path = docx_path.rsplit(".", 1)[0] + ".pdf"
        if not os.path.exists(pdf_path):
            raise RuntimeError(f"Geen PDF geproduceerd. stdout: {result.stdout}, stderr: {result.stderr}")

        with open(pdf_path, "rb") as f:
            return f.read()


@app.post("/generate-pdf")
async def generate_pdf(data: GenerateRequest):
    builder = TEMPLATE_BUILDERS.get(data.template or "1")
    if builder is None:
        raise HTTPException(status_code=400, detail=f"Onbekend template: {data.template}")

    docx_bytes = builder(data)
    try:
        pdf_bytes = convert_docx_to_pdf(docx_bytes)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

    return StreamingResponse(
        _io.BytesIO(pdf_bytes),
        media_type="application/pdf",
        headers={"Content-Disposition": 'attachment; filename="offerte.pdf"'},
    )
