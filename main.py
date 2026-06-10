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
# STEP 1 — STRUCTURE EXTRACTION (NOW JSON = MUCH STRONGER)
# =========================================================
def extract_blueprint(example_text):

    SYSTEM = """
You are a DOCUMENT STRUCTURE ENGINE.

Extract the document into STRICT JSON format.

RULES:
- No explanations
- No extra text
- Output VALID JSON ONLY

FORMAT:

{
  "sections": [
    {
      "title": "",
      "type": "text | bullets | table",
      "subsections": [
        {
          "title": "",
          "type": "text | bullets | table",
          "table_columns": []
        }
      ]
    }
  ]
}

IMPORTANT:
- Preserve exact order
- Detect tables precisely
- Detect bullet sections
- Detect headings
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
# STEP 2 — CONTENT GENERATION PER SECTION (LOCKED SYSTEM)
# =========================================================
def generate_section(section, transcript, notes, previous_consult):

    SYSTEM = """
You fill in a SINGLE section of a medical document.

RULES:
- Follow section type EXACTLY
- Do not change structure
- Be extremely detailed
- No summarization
- Medical precision required
- Output only section content
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

Fill ONLY this section with correct content.
"""

    response = client.chat.completions.create(
        model="gpt-5-mini",
        messages=[
            {"role": "system", "content": SYSTEM},
            {"role": "user", "content": USER}
        ],
        temperature=1
    )

    return response.choices[0].message.content


# =========================================================
# STEP 3 — FULL DOCUMENT GENERATION
# =========================================================
def generate_document(transcript, notes, example_text, previous_consult=""):

    blueprint = extract_blueprint(example_text)

    final_sections = []

    for section in blueprint["sections"]:
        filled = generate_section(section, transcript, notes, previous_consult)
        final_sections.append({
            "title": section["title"],
            "content": filled,
            "type": section["type"]
        })

    return final_sections


# =========================================================
# STEP 4 — WORD EXPORT (PROPER STRUCTURE)
# =========================================================
def generate_word(sections, output_file):

    doc = Document()

    for sec in sections:

        # Heading
        doc.add_heading(sec["title"], level=1)

        content = sec["content"]

        for line in content.split("\n"):

            line = line.strip()
            if not line:
                continue

            # TABLE DETECTION
            if "|" in line:
                cells = [c.strip() for c in line.split("|")]
                table = doc.add_table(rows=1, cols=len(cells))
                for i, cell in enumerate(cells):
                    table.rows[0].cells[i].text = cell
                continue

            # BULLETS
            if line.startswith("•"):
                doc.add_paragraph(line, style="List Bullet")
            else:
                doc.add_paragraph(line)

    doc.save(output_file)
    return output_file