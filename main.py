import os
import json
from faster_whisper import WhisperModel
from docxtpl import DocxTemplate
from docx import Document

# =============================
# INIT
# =============================
import streamlit as st
from openai import OpenAI

client = OpenAI(
    api_key=st.secrets["OPENAI_API_KEY"]
)

@st.cache_resource
def load_whisper():
    return WhisperModel("base")

whisper_model = load_whisper()

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
BASISDOCUMENT = parse_basisdocument(read_docx("basis.docx"))
BASISDOCUMENT_TEXT = "\n".join(BASISDOCUMENT)

# voorbeelden tijdelijk uitgeschakeld
# VOORBEELD1 = read_docx("voorbeeld1.docx")
# VOORBEELD2 = read_docx("voorbeeld2.docx")
# VOORBEELD3 = read_docx("voorbeeld3.docx")

# =============================
# AUDIO → TEXT
# =============================
def transcribe_audio(file_path):
    segments, _ = whisper_model.transcribe(file_path)
    return " ".join([s.text for s in segments]).strip()

# =============================
# AI → JSON OUTPUT
# =============================
def generate_json(transcript, notes=""):

    SYSTEM = f"""
    DOEL:
    Schrijf een praktisch en concreet consultverslag voor de cliënt zelf.
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
    - Duidelijk, praktisch en volledig
    - Geen algemene adviezen
    - Geen vage formuleringen
    - Geen herhaling
    - Formeel Nederlands
    - Schrijf duidelijk en volledig
    - Het verslag mag uitgebreid zijn wanneer nodig
    - Vermijd onnodige herhaling of algemene uitleg
    - Minimaal 500 woorden, geen maximum
    - Volledigheid en duidelijkheid zijn belangrijker dan lengte
    - Spreek de cliënt aan met "u"
    - Gebruik formele aanspreekvormen
    - Gebruik nooit "je" of "jij"

    BELANGRIJKE REGELS:
    - Geef ALTIJD geldige JSON
    - Geen tekst buiten JSON
    - Geen uitleg
    - Geen hallucinaties
    - Geen verwijzingen naar transcript, basisdocument of welke bron dan ook    
    - Geen woorden zoals: "volgens", "bron", "uit transcript", "gebaseerd op", "advies"
    - Gebruik transcript en notities als hoofdbron
    - Gebruik BASISDOCUMENT alleen om supplementdetails aan te vullen
    - Voeg NOOIT supplementen toe die niet genoemd zijn
    - Schrijf nooit:
    "niet vermeld in transcript"
    "niet gevonden in basisdocument"
    of vergelijkbare uitleg

    INHOUD:
    - Schrijf alsof dit direct naar de cliënt gestuurd wordt
    - Wees specifiek over klachten en acties
    - Benoem concrete voeding die verminderd/verhoogd moet worden
    - Benoem concrete supplement-instructies
    - Benoem concrete vervolgstappen

    BULLETS:
    - "huidige_situatie"
    - "voeding_verminderen"
    - "voeding_verhogen"
    - "onderzoeken"
    - "therapeuten"

    moeten:
    - korte bullets zijn
    - elke bullet starten met "•"
    - elke bullet op nieuwe regel

    SUPPLEMENT SEARCH RULE:
    - Scan BASISDOCUMENT regel voor regel
    - Match op woorden (niet exacte naam)
    - Magnesium bisglycinaat = match "magnesium" + "bisglycinaat"
    - Geen filtering door model toegestaan
    - Neem alleen relevante details over
    - Gebruik korte concrete bullets
    - Vermeld:
    - dosering
    - gebruiksmoment
    - opbouw
    - prijs
    - houdbaarheid
    indien beschikbaar

    MOMENTEN:
    - "voor ontbijt" =
    "voor_ontbijt": true
    "ontbijt": false

    - "bij ontbijt" =
    "ontbijt": true
    "voor_ontbijt": false

    - Zet nooit beide op true tenzij expliciet genoemd

    ALS INFORMATIE ONTBREEKT:
    - Gebruik exact: "Onbekend"
    - Gebruik NOOIT "NVT"
    - Bij supplementdetails:
    - BIJ supplementen:
    ontbrekende velden volledig weglaten (NIET invullen met "Onbekend")

    CONTROLE:
    Controleer VOOR output of ALLE genoemde:
    - klachten
    - supplementen
    - acties
    - voeding
    - onderzoeken
    - symptomen
    zijn verwerkt.

    Verwijder niets.
    
    ====================
    BASISDOCUMENT:
    {BASISDOCUMENT_TEXT}

    ====================
    JSON:

    {{
    "datum": "",
    "naam": "",
    "volgende_consult": "",
    "huidige_situatie": "",
    "voeding_verminderen": "",
    "voeding_verhogen": "",
    "onderzoeken": "",
    "therapeuten": "",
    "supplementen": [
        {{
        "naam": "",
        "details": [],
        "voor_ontbijt": false,
        "ontbijt": false,
        "tussen_1": false,
        "lunch": false,
        "tussen_2": false,
        "diner": false,
        "voor_slapen": false
        }}
    ]
    }} 
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
            {"role": "system", "content": SYSTEM},
            {"role": "user", "content": USER}
        ],
        temperature=1,
        response_format={"type": "json_object"}
    )

    return json.loads(response.choices[0].message.content)

def clean_supplements(data):
    for s in data.get("supplementen", []):
        # verwijder lege of "Onbekend" details
        if "details" in s:
            s["details"] = [
                d for d in s["details"]
                if d and d.strip() and d != "Onbekend"
            ]

            # als leeg → volledig weghalen
            if not s["details"]:
                s["details"] = []
    return data

def strip_bullets(data):
    for s in data.get("supplementen", []):
        if "details" in s:
            s["details"] = [
                d.replace("•", "").strip() for d in s["details"]
                if d
            ]
    return data

def kruis(val):
    return "✖" if val is True else ""
    
# =============================
# WORD GENERATOR
# =============================
def generate_word(data, output="verslag.docx"):
    for s in data["supplementen"]:
        s["kruis_voor_ontbijt"] = kruis(s["voor_ontbijt"])
        s["kruis_ontbijt"] = kruis(s["ontbijt"])
        s["kruis_tussen_1"] = kruis(s["tussen_1"])
        s["kruis_lunch"] = kruis(s["lunch"])
        s["kruis_tussen_2"] = kruis(s["tussen_2"])
        s["kruis_diner"] = kruis(s["diner"])
        s["kruis_voor_slapen"] = kruis(s["voor_slapen"])

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
        transcript = input("Plak transcript:\n")
    elif keuze == "2":
        pad = input("Audio pad: ")
        transcript = transcribe_audio(pad)
        print("\nTranscript geladen.\n")
    else:
        print("Ongeldige keuze")
        exit()

    notes = input("Notities:\n")

    print("\nAI bezig...\n")

    data = generate_json(transcript, notes)
    data = clean_supplements(data)
    data = strip_bullets(data)

    print("\nWord genereren...\n")

    generate_word(data)

    print("\nKLAAR")