import os
import json
from dotenv import load_dotenv
from openai import OpenAI
from faster_whisper import WhisperModel
from docxtpl import DocxTemplate
from docx import Document
from rapidfuzz import fuzz

# =============================
# INIT
# =============================
load_dotenv()

client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

if not os.getenv("OPENAI_API_KEY"):
    raise ValueError("API key ontbreekt")

whisper_model = WhisperModel("base")

# =============================
# HELPERS
# =============================
def read_docx(file_path):
    doc = Document(file_path)
    return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])

def parse_basisdocument(text):
    items = []

    for line in text.split("\n"):
        line = line.strip()

        if line:
            items.append(line)

    return items

# =============================
# LOAD STATIC FILES
# =============================
BASISDOCUMENT = parse_basisdocument(
    read_docx("basis.docx")
)

# =============================
# AUDIO → TEXT
# =============================
def transcribe_audio(file_path):

    segments, _ = whisper_model.transcribe(file_path)

    transcript_parts = []

    for s in segments:
        transcript_parts.append(
            f"[{round(s.start)}s] {s.text}"
        )

    return "\n".join(transcript_parts).strip()

# =============================
# EXTRACT SUPPLEMENTS
# =============================
def extract_supplements(transcript, notes=""):

    SYSTEM = """
    Extract ALL supplements, medications, vitamins, minerals, herbs, probiotics and health products mentioned.

    ALSO EXTRACT:
    - dosage
    - timing (morning / lunch / evening / sleep)
    - price (if mentioned)
    - brand
    - usage instructions
    - build-up schedule (if mentioned)

    Return ONLY valid JSON.

    STRICT FORMAT:

    {
    "supplementen": [
        {
        "naam": "",
        "merk": "",
        "dosering": "",
        "moment": "",
        "prijs": "",
        "gebruik": "",
        "opbouw": ""
        }
    ]
    }

    RULES:
    - If unknown use empty string ""
    - Do NOT invent anything
    - Only extract from transcript
    - No extra text outside JSON
    """

    USER = f"""
    TRANSCRIPT:
    {transcript}

    NOTITIES:
    {notes}
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
                "content": USER
            }
        ],
        temperature=1,
        response_format={
            "type": "json_object"
        }
    )

    data = json.loads(
        response.choices[0].message.content
    )

    return data.get("supplementen", [])

# =============================
# SEARCH BASISDOCUMENT
# =============================
def search_basisdocument(supplements):

    results = {}

    for supp in supplements:

        # FIX: pak alleen naam
        name = supp.get("naam", "")

        matches = []

        for line in BASISDOCUMENT:

            score = fuzz.partial_ratio(
                name.lower(),
                line.lower()
            )

            if score >= 80:
                matches.append(line)

        results[name] = matches

    return results
# =============================
# AI → JSON OUTPUT
# =============================
def generate_json(
    transcript,
    notes="",
    supplement_info=None
):

    if supplement_info is None:
        supplement_info = {}

    SYSTEM = """
    DOEL:
    Schrijf een praktisch en concreet
    consultverslag voor de cliënt zelf.

    De cliënt moet exact begrijpen:
    - wat er besproken is
    - wat de oorzaak/problemen zijn
    - wat hij/zij moet doen
    - wanneer supplementen genomen moeten worden
    - wat vermeden of verhoogd moet worden
    - welke onderzoeken nog moeten gebeuren

    SCHRIJFSTIJL:
    - Concreet
    - Praktisch
    - Duidelijk
    - Geen algemene adviezen
    - Geen vage formuleringen
    - Geen herhaling
    - Formeel Nederlands
    - Spreek de cliënt aan met "u"

    BELANGRIJKE REGELS:
    - Geef ALTIJD geldige JSON
    - Geen tekst buiten JSON
    - Geen uitleg
    - Geen hallucinaties
    - Gebruik transcript als hoofdbron
    - Gebruik supplement database alleen
      voor aanvulling van details
    - Voeg NOOIT supplementen toe
      die niet genoemd zijn

    CRITICAL:
    Voor afronding:
    - Controleer of ALLE supplementen
      uit transcript aanwezig zijn
    - Controleer of ALLE symptomen aanwezig zijn
    - Controleer of ALLE acties aanwezig zijn
    - Controleer of ALLE timing instructies aanwezig zijn

    BULLETS:
    - huidige_situatie
    - voeding_verminderen
    - voeding_verhogen
    - onderzoeken
    - therapeuten

    moeten:
    - korte bullets zijn
    - starten met "•"
    - elke bullet op nieuwe regel

    MOMENTEN:
    - voor ontbijt =
      voor_ontbijt true

    - bij ontbijt =
      ontbijt true

    JSON:

    {
        "datum": "",
        "naam": "",
        "volgende_consult": "",
        "huidige_situatie": "",
        "voeding_verminderen": "",
        "voeding_verhogen": "",
        "onderzoeken": "",
        "therapeuten": "",
        "supplementen": [
            {
                "naam": "",
                "details": [],
                "voor_ontbijt": false,
                "ontbijt": false,
                "tussen_1": false,
                "lunch": false,
                "tussen_2": false,
                "diner": false,
                "voor_slapen": false
            }
        ]
    }
    """

    USER = f"""
    TRANSCRIPT:
    {transcript}

    SUPPLEMENT DATABASE RESULTS:
    {json.dumps(supplement_info, ensure_ascii=False)}

    NOTITIES:
    {notes}
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
                "content": USER
            }
        ],
        temperature=1,
        response_format={
            "type": "json_object"
        }
    )

    return json.loads(
        response.choices[0].message.content
    )

# =============================
# CLEANING
# =============================
def clean_supplements(data):

    for s in data.get("supplementen", []):

        if "details" in s:

            s["details"] = [
                d for d in s["details"]
                if d and d.strip()
                and d != "Onbekend"
            ]

            if not s["details"]:
                s["details"] = []

    return data

def strip_bullets(data):

    for s in data.get("supplementen", []):

        if "details" in s:

            s["details"] = [
                d.replace("•", "").strip()
                for d in s["details"]
                if d
            ]

    return data

def kruis(val):
    return "✖" if val is True else ""

# =============================
# WORD GENERATOR
# =============================
def generate_word(
    data,
    output="verslag.docx"
):

    for s in data["supplementen"]:

        s["kruis_voor_ontbijt"] = kruis(
            s["voor_ontbijt"]
        )

        s["kruis_ontbijt"] = kruis(
            s["ontbijt"]
        )

        s["kruis_tussen_1"] = kruis(
            s["tussen_1"]
        )

        s["kruis_lunch"] = kruis(
            s["lunch"]
        )

        s["kruis_tussen_2"] = kruis(
            s["tussen_2"]
        )

        s["kruis_diner"] = kruis(
            s["diner"]
        )

        s["kruis_voor_slapen"] = kruis(
            s["voor_slapen"]
        )

    doc = DocxTemplate("template.docx")

    doc.render(data)

    doc.save(output)

    print(f"Word opgeslagen: {output}")

# =============================
# MAIN
# =============================
if __name__ == "__main__":

    print("=== AI Consult Generator ===")

    keuze = input("1 = tekst | 2 = audio: ")

    if keuze == "1":

        transcript = input(
            "Plak transcript:\n"
        )

    elif keuze == "2":

        pad = input("Audio pad: ")

        transcript = transcribe_audio(pad)

        print("\nTranscript geladen.\n")

    else:
        print("Ongeldige keuze")
        exit()

    notes = input("Notities:\n")

    print("\nSupplementen analyseren...\n")

    supplements = extract_supplements(
        transcript,
        notes
    )

    supplement_info = search_basisdocument(
        supplements
    )

    print("\nAI verslag genereren...\n")

    data = generate_json(
        transcript,
        notes,
        supplement_info
    )

    data = clean_supplements(data)

    data = strip_bullets(data)

    print("\nWord genereren...\n")

    generate_word(data)

    print("\nKLAAR")