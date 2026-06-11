import streamlit as st
import uuid
from faster_whisper import WhisperModel
from main import (
    read_docx,
    extract_style_dna,
    extract_clinical_facts,
    compose_document,
    generate_word
)

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
# UI
# =========================
st.title("🧠 Clinical Report Generator (Advanced AI)")

example_file = st.file_uploader("Upload example DOCX (style reference)", type=["docx"])
audio_file = st.file_uploader("Upload audio", type=["mp3", "wav"])
notes = st.text_area("Extra notes")


if st.button("Generate Document"):

    if not example_file:
        st.error("Upload example document")
        st.stop()

    # =========================
    # INPUT PROCESSING
    # =========================
    example_text = read_docx(example_file)

    transcript = ""
    if audio_file:
        temp_path = "temp_audio.wav"
        with open(temp_path, "wb") as f:
            f.write(audio_file.read())

        transcript = transcribe_audio(temp_path)

    # =========================
    # PIPELINE
    # =========================
    with st.spinner("Extracting style DNA..."):
        style_dna = extract_style_dna(example_text)

    with st.spinner("Extracting clinical facts..."):
        facts = extract_clinical_facts(transcript, notes)

    with st.spinner("Composing document..."):
        structured_doc = compose_document(style_dna, facts)

    with st.spinner("Generating Word file..."):
        output_file = f"report_{uuid.uuid4().hex}.docx"
        generate_word(structured_doc, output_file)

    st.success("Done!")

    with open(output_file, "rb") as f:
        st.download_button(
            "Download Word Document",
            f,
            file_name="clinical_report.docx"
        )