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

        TUSSEN_MOMENTEN:
        - tussen_1 = tussen ontbijt en lunch
        - tussen_2 = tussen lunch en diner

        - Zet nooit beide op true tenzij expliciet genoemd

        ALS INFORMATIE ONTBREEKT:
        - Gebruik exact: "Onbekend"
        - Gebruik NOOIT "NVT"
        - Bij supplementdetails:
        - BIJ supplementen:
        ontbrekende velden volledig weglaten (NIET invullen met "Onbekend")

        ==================================================
        UITGEBREIDE VELDSPECIFICATIES
        ==================================================

        De onderstaande regels hebben voorrang op algemene schrijfinstructies wanneer zij specifieker zijn.

        Alle informatie moet uitsluitend worden gehaald uit:
        1. Transcript
        2. Notities
        3. Basisdocument (alleen voor supplementdetails)

        Voeg nooit informatie toe die niet expliciet aanwezig is.

        ==================================================
        HUIDIGE_SITUATIE
        ==================================================

        Schrijf een overzicht van de huidige situatie op basis van observaties uit transcript en notities.

        Doel:
        Beschrijven wat er de afgelopen periode is gebeurd, welke acties zijn uitgevoerd, welke veranderingen zijn opgetreden en hoe dit verlopen is.

        Regels:
        - Observeer en beschrijf
        - Geen interpretaties die niet genoemd zijn
        - Geen aannames
        - Concreet en volledig
        - Beschrijf voortgang, klachten, reacties, energie, gewicht, slaap, ontlasting, behandelingstrajecten en andere relevante ontwikkelingen
        - Gebruik meerdere bullets indien nodig
        - Elke bullet start met "•"

        ==================================================
        VOEDING_VERMINDEREN
        ==================================================

        Vermeld alle voedingsmiddelen, voedingsgroepen of voedingsgewoonten die verminderd, vermeden, gestopt of aangepast moeten worden.

        Per onderdeel vermelden indien aanwezig:

        - Wat verminderd of vermeden moet worden
        - Concrete voorbeelden van producten
        - Duur van het advies
        - Wat te doen bij afwijking
        - Welke reacties geobserveerd moeten worden
        - Eventuele aandachtspunten rond gewicht, energie, darmfunctie, huid of andere klachten

        Behoud ook eerder gegeven adviezen indien expliciet vermeld als nog lopend.

        Gebruik bullets.

        ==================================================
        VOEDING_VERHOGEN
        ==================================================

        Vermeld alle voedingsmiddelen, voedingsgroepen of voedingsgewoonten die actief verhoogd of toegevoegd moeten worden.

        Per onderdeel vermelden indien aanwezig:

        - Wat verhoogd moet worden
        - Praktische toepassing
        - Voorbeelden
        - Doel of gewenste ondersteuning
        - Aandachtspunten voor observatie

        Gebruik bullets.

        ==================================================
        ONDERZOEKEN
        ==================================================

        Vermeld uitsluitend onderzoeken, testen, metingen of controles die daadwerkelijk genoemd worden.

        Per onderzoek vermelden indien aanwezig:

        - Wat onderzocht wordt
        - Wie het uitvoert
        - Of een testset besteld moet worden
        - Relevante bloedwaarden
        - Relevante meetwaarden
        - Praktische vervolgstappen

        Gebruik bullets.

        Indien geen onderzoeken genoemd worden:
        gebruik een lege string.

        ==================================================
        THERAPEUTEN
        ==================================================

        Vermeld uitsluitend behandelaars die expliciet genoemd worden.

        Per behandelaar:

        - Naam of functie
        - Behandeling
        - Status of voortgang

        Gebruik bullets.

        Indien niet genoemd:
        gebruik een lege string.

        ==================================================
        SUPPLEMENTEN
        ==================================================

        Voor ieder genoemd supplement:

        Zoek in het volledige BASISDOCUMENT naar alle relevante informatie.

        Neem alleen informatie over die daadwerkelijk gevonden wordt.

        Gebruik afzonderlijke bullets voor iedere detailregel.

        Zoek indien aanwezig naar:

        - ingrediënten
        - dosering
        - gebruiksmoment
        - opbouwschema (de verschillende stappen in sub bullets)
        - inhoud verpakking
        - aantal capsules
        - aantal tabletten
        - aantal druppels
        - prijs
        - bewaarinstructies
        - waarschuwingen
        - relevante toepassing

        Prijs altijd opnemen indien beschikbaar.

        Nooit meerdere details combineren in één bullet.

        ==================================================
        SUPPLEMENT-INNAME
        ==================================================

        Vul tijdsvakken zo nauwkeurig mogelijk.

        Gebruik exacte hoeveelheden.

        Voorbeelden:

        (1 cap) of 1 = "1 capsule"
        (2 cap) of 2 = "2 capsules"
        (1 dr) = "1 druppel"
        (2 dr) = "2 druppels"
        (om de dag)

        Gebruik uitsluitend informatie die daadwerkelijk genoemd wordt.

        ==================================================
        MINERALENOVERZICHT
        ==================================================

        Analyseer transcript en notities volledig.

        Zoek naar:

        - mineralen
        - druppels
        - mineraalsupplementen
        - doseringswijzigingen

        Indien voldoende informatie aanwezig is, maak aan het einde van "huidige_situatie" een apart overzicht met:

        - mineraal
        - huidige dosering
        - eerste doel
        - tweede doel
        - opmerkingen

        Gebruik uitsluitend letterlijk aanwezige informatie.

        Indien geen mineralen genoemd worden:
        geen overzicht opnemen.

        MINERALENOVERZICHT REGEL:
        - Het mineralenoverzicht wordt ALTIJD als LAATSTE veld in de JSON geplaatst.
        - Het staat dus onderaan het JSON-object, direct vóór de afsluitende.
        - Het mag niet in een ander veld staan (zoals huidige_situatie).
        - Het mag niet buiten de JSON worden geplaatst.
        - Als er geen mineralen zijn: veld blijft leeg ("").

        ==================================================
        BEHOUD VAN VORIGE ADVIEZEN
        ==================================================

        Wanneer transcript of notities aangeven dat eerdere adviezen nog steeds gelden:

        - behoud deze adviezen
        - neem ze opnieuw op in het relevante onderdeel
        - verwijder ze niet

        ==================================================
        EINDCONTROLE
        ==================================================

        Controleer vóór output dat alle genoemde:

        - klachten
        - symptomen
        - voeding
        - acties
        - onderzoeken
        - therapeuten
        - supplementen
        - doseringswijzigingen
        - vervolgafspraken

        volledig zijn verwerkt.

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