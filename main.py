import uuid
import json
from docx import Document
from faster_whisper import WhisperModel
from openai import OpenAI
import streamlit as st

# =========================
# OPENAI
# =========================
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

# =========================
# WHISPER
# =========================
@st.cache_resource
def load_whisper():
    return WhisperModel("base")

whisper_model = load_whisper()

def transcribe_audio(file_path):
    segments, _ = whisper_model.transcribe(file_path)
    return " ".join([s.text for s in segments]).strip()

# =========================
# DOCX READER
# =========================
def read_docx(file):
    doc = Document(file)
    return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])

# =========================================================
# STEP 1 — STRUCTURE EXTRACTION (STRONG JSON)
# =========================================================
def extract_blueprint(example_text):

    SYSTEM = """
You are a CLINICAL DOCUMENT STRUCTURE ENGINE.

Extract STRICT JSON blueprint.

RULES:
- Preserve order
- Detect sections exactly
- Detect tables, bullets, paragraphs
- NO content rewriting
- OUTPUT ONLY JSON

FORMAT:

{
  "sections": [
    {
      "title": "",
      "blocks": [
        {
          "type": "paragraph | bullets | table",
          "content": "",
          "items": [],
          "columns": [],
          "rows": []
        }
      ]
    }
  ]
}
"""

    response = client.chat.completions.create(
        model="gpt-5-mini",
        messages=[
            {"role": "system", "content": SYSTEM},
            {"role": "user", "content": example_text}
        ],
        temperature=1,
        response_format={"type": "json_object"}
    )

    return json.loads(response.choices[0].message.content)

# =========================================================
# STEP 2 — FILL SECTION (BLOCK BASED)
# =========================================================
def generate_section(section, transcript, notes, previous_consult):

    SYSTEM = """
You are a MEDICAL DOCUMENT FILL ENGINE.

RULES:
- Follow structure EXACTLY
- Fill each block correctly
- No summarization
- No merging blocks
- Keep bullets atomic (1 idea per bullet)
- Output VALID JSON for blocks only
"""

    USER = f"""
SECTION STRUCTURE:
{json.dumps(section, indent=2)}

TRANSCRIPT:
{transcript}

NOTES:
{notes}

PREVIOUS:
{previous_consult}

Fill this section with correct clinical content.
Return same structure, only filled content.
"""

    response = client.chat.completions.create(
        model="gpt-5-mini",
        messages=[
            {"role": "system", "content": SYSTEM},
            {"role": "user", "content": USER}
        ],
        temperature=1,
        response_format={"type": "json_object"}
    )

    return json.loads(response.choices[0].message.content)

# =========================================================
# STEP 3 — FULL DOCUMENT
# =========================================================
def generate_document(transcript, notes, example_text, previous_consult=""):

    blueprint = extract_blueprint(example_text)

    final = []

    for section in blueprint["sections"]:
        filled = generate_section(section, transcript, notes, previous_consult)
        final.append(filled)

    return final

# =========================================================
# STEP 4 — WORD EXPORT (PROFESSIONAL LAYOUT)
# =========================================================
def generate_word(sections, output_file):

    doc = Document()

    for sec in sections:

        doc.add_heading(sec["title"], level=1)

        for block in sec["blocks"]:

            btype = block.get("type")

            # PARAGRAPH
            if btype == "paragraph":
                doc.add_paragraph(block.get("content", ""))

            # BULLETS
            elif btype == "bullets":
                for item in block.get("items", []):
                    doc.add_paragraph(item, style="List Bullet")

            # TABLE
            elif btype == "table":
                cols = block.get("columns", [])
                rows = block.get("rows", [])

                if cols:
                    table = doc.add_table(rows=1, cols=len(cols))

                    for i, c in enumerate(cols):
                        table.rows[0].cells[i].text = str(c)

                    for r in rows:
                        row_cells = table.add_row().cells
                        for i, val in enumerate(r):
                            row_cells[i].text = str(val)

    doc.save(output_file)
    return output_file