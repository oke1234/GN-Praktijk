import json
from docx import Document
from openai import OpenAI
import streamlit as st

# =========================
# OPENAI CLIENT
# =========================
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])


# =========================
# DOCX READER
# =========================
def read_docx(file):
    doc = Document(file)
    return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])


# =========================
# STEP 1 — STYLE DNA EXTRACTION
# =========================
def extract_style_dna(example_text):
    """
    Extracts structure + formatting logic from example document.
    """

    SYSTEM = """
You are a CLINICAL DOCUMENT STYLE ANALYZER.

Extract a STYLE DNA in JSON.

Focus on:
- section order
- naming patterns
- use of tables
- bullet density
- writing style (formal/clinical/structured)
- repetition patterns
- hierarchy rules

RULES:
- OUTPUT VALID JSON ONLY
- NO commentary
"""

    response = client.chat.completions.create(
        model="gpt-5-mini",
        messages=[
            {"role": "system", "content": SYSTEM},
            {"role": "user", "content": example_text}
        ],
        response_format={"type": "json_object"},
        temperature=1
    )

    return json.loads(response.choices[0].message.content)


# =========================
# STEP 2 — CLINICAL FACT EXTRACTION
# =========================
def extract_clinical_facts(transcript, notes):
    """
    Converts messy input into structured medical facts.
    """

    SYSTEM = """
You are a CLINICAL FACT EXTRACTION ENGINE.

Convert input into structured medical facts.

Return JSON:

{
  "symptoms": [],
  "observations": [],
  "supplements": [],
  "actions": [],
  "risks": [],
  "timeline_events": []
}

RULES:
- No storytelling
- No full sentences unless needed
- Be precise
- Extract everything important
"""

    response = client.chat.completions.create(
        model="gpt-5-mini",
        messages=[
            {"role": "system", "content": SYSTEM},
            {"role": "user", "content": f"TRANSCRIPT:\n{transcript}\n\nNOTES:\n{notes}"}
        ],
        response_format={"type": "json_object"},
        temperature=1
    )

    return json.loads(response.choices[0].message.content)

# =========================
# STEP 3 — DOCUMENT COMPOSER
# =========================
def compose_document(style_dna, clinical_facts):
    """
    Builds structured document using style + facts.
    """

    SYSTEM = """
You are a CLINICAL DOCUMENT COMPOSER.

You MUST:
- follow STYLE DNA exactly
- convert clinical facts into structured sections
- ensure clarity + hierarchy
- NEVER output free text
- ALWAYS output JSON

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
            {
                "role": "system",
                "content": SYSTEM
            },
            {
                "role": "user",
                "content": json.dumps({
                    "style_dna": style_dna,
                    "clinical_facts": clinical_facts
                })
            }
        ],
        response_format={"type": "json_object"},
        temperature=1
    )

    return json.loads(response.choices[0].message.content)

# =========================
# STEP 4 — WORD EXPORT ENGINE
# =========================
def generate_word(structured_doc, output_file):

    doc = Document()

    for section in structured_doc["sections"]:

        # BIG TITLE
        doc.add_heading(section["title"], level=1)

        for block in section.get("blocks", []):

            btype = block.get("type")

            # =========================
            # PARAGRAPH BLOCK
            # =========================
            if btype == "paragraph":
                text = block.get("content", "")
                if text:
                    doc.add_paragraph(text)

            # =========================
            # BULLET BLOCK
            # =========================
            elif btype == "bullets":
                for item in block.get("items", []):
                    doc.add_paragraph(item, style="List Bullet")

            # =========================
            # TABLE BLOCK
            # =========================
            elif btype == "table":
                cols = block.get("columns", [])
                rows = block.get("rows", [])

                if cols:
                    table = doc.add_table(rows=1, cols=len(cols))

                    # header
                    for i, col in enumerate(cols):
                        table.rows[0].cells[i].text = str(col)

                    # rows
                    for r in rows:
                        row_cells = table.add_row().cells
                        for i, val in enumerate(r):
                            if i < len(row_cells):
                                row_cells[i].text = str(val)

    doc.save(output_file)
    return output_file