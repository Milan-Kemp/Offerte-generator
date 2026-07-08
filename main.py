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
