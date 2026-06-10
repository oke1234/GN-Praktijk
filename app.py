import streamlit as st
import uuid

from main import (
    read_docx,
    generate_document,
    generate_word,
    transcribe_audio
)

# =========================
# UI
# =========================
st.title("AI Consult Verslag Generator (Style Clone)")

# =========================
# EXAMPLE DOCUMENT (IMPORTANT)
# =========================
st.subheader("📄 Voorbeeld (stijl template)")

example_file = st.file_uploader(
    "Upload voorbeeld document (DOCX)",
    type=["docx"]
)

example_text = ""

if example_file:
    example_text = read_docx(example_file)
    st.success("Voorbeeld geladen")

# =========================
# VORIG CONSULT (OPTIONAL)
# =========================
st.subheader("📋 Vorig consult (optioneel)")

previous_file = st.file_uploader(
    "Upload vorig consult (DOCX)",
    type=["docx"],
    key="prev"
)

previous_consult = ""

if previous_file:
    previous_consult = read_docx(previous_file)

# =========================
# TRANSCRIPT
# =========================
st.subheader("🎤 Transcript")

audio_file = st.file_uploader("Upload audio (mp3/wav)", type=["mp3", "wav"])

transcript = ""

if audio_file:
    with open("temp_audio.mp3", "wb") as f:
        f.write(audio_file.read())

    transcript = transcribe_audio("temp_audio.mp3")
    st.success("Transcript gemaakt")

else:
    transcript = st.text_area("Of plak transcript")

# =========================
# NOTES
# =========================
st.subheader("📝 Notities")
notes = st.text_area("Extra notities")

# =========================
# GENERATE
# =========================
if st.button("Genereer document"):

    if not example_text:
        st.error("Upload eerst een voorbeeld document")
        st.stop()

    if not transcript:
        st.error("Geen transcript")
        st.stop()

    with st.spinner("AI analyseert stijl en maakt document..."):

        output_text = generate_document(
            transcript=transcript,
            notes=notes,
            example_text=example_text,
            previous_consult=previous_consult
        )

        output_file = f"verslag_{uuid.uuid4()}.docx"
        generate_word(output_text, output_file)

    st.success("Klaar!")

    with open(output_file, "rb") as f:
        st.download_button(
            "Download Word document",
            f,
            file_name="consult_verslag.docx"
        )