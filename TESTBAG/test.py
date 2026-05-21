from faster_whisper import WhisperModel

model = WhisperModel("base")

def transcribe(file_path):
    segments, info = model.transcribe(file_path)

    text = ""
    for seg in segments:
        text += seg.text + " "

    return text


if __name__ == "__main__":
    path = input("Pad naar audio: ")
    print(transcribe(path))