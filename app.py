import streamlit as st
from docx import Document
from main import generate_word, generate_json, clean_supplements, strip_bullets, transcribe_audio

# ---------- UI ----------
st.title("Consult Verslag Generator")

# ---------- TRANSCRIPT ----------
st.subheader("Transcript")

audio_file = st.file_uploader("Upload audio (mp3/wav)", type=["mp3", "wav"])

transcript = ""

if audio_file:
    with open("temp_audio.mp3", "wb") as f:
        f.write(audio_file.read())

    transcript = transcribe_audio("temp_audio.mp3")
    st.success("Audio getranscribeerd")

else:
    transcript = st.text_area("Of plak transcript")

# ---------- NOTITIES ----------
st.subheader("Notities")
notes = st.text_area("Typ notities hier")

# ---------- GENERATE ----------
if st.button("Genereer Word document"):

    if not transcript:
        st.error("Geen transcript gevonden")
        st.stop()

    with st.spinner("AI is bezig..."):

        data = generate_json(transcript, notes)
        data = clean_supplements(data)
        data = strip_bullets(data)

        import uuid
        output_file = f"verslag_{uuid.uuid4()}.docx"
        
        generate_word(data, output_file)

    st.success("Klaar!")

    with open(output_file, "rb") as f:
        st.download_button(
            "Download Word document",
            f,
            file_name="consult_verslag.docx"
        )