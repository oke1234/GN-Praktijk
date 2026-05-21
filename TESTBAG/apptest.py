import streamlit as st
from faster_whisper import WhisperModel

model = WhisperModel("base") 

def transcribe(file):
    segments, info = model.transcribe(file)

    text = ""
    for seg in segments:
        text += seg.text + " "

    return text


st.title("Free Whisper Test (offline)")

audio = st.file_uploader("Upload mp3/wav")

if audio:
    with open("temp.mp3", "wb") as f:
        f.write(audio.read())

    if st.button("Transcribe"):
        result = transcribe("temp.mp3")
        st.write(result)