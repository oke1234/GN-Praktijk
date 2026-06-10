import json
import uuid
from docx import Document
from docxtpl import DocxTemplate
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
# READ DOCX
# =========================
def read_docx(file):
    doc = Document(file)
    return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])

# =========================
# AI CORE (STYLE TRANSFER)
# =========================
def generate_document(transcript, notes, example_text, previous_consult=""):

    SYSTEM = """
You are a professional document style replication AI.

TASK:
1. Analyze the EXAMPLE document.
2. Extract:
   - structure (sections, order)
   - writing tone
   - bullet style
   - formatting patterns
   - level of detail
3. Apply EXACT same structure and style to new input.

RULES:
- Do NOT copy content from example
- Only copy style + structure
- Keep formatting identical
- Do NOT output JSON
- Output clean formatted text only
- Do NOT explain anything
"""

    USER = f"""
EXAMPLE DOCUMENT (STYLE TEMPLATE):
{example_text}

PREVIOUS CONSULT (optional context):
{previous_consult}

NEW TRANSCRIPT:
{transcript}

NOTES:
{notes}

Generate a new document using the SAME style and structure as the example.
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


# =========================
# WORD EXPORT
# =========================
def generate_word(text, output_file):
    doc = Document()
    for line in text.split("\n"):
        doc.add_paragraph(line)
    doc.save(output_file)
    return output_file