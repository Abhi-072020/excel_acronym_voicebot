import pandas as pd
import speech_recognition as sr
import pyttsx3
import time

# Load acronyms from Excel
def load_acronyms(file_path):
    df = pd.read_excel(file_path)
    return dict(zip(df['Acronym'], df['Full Form']))


# Speech to Text
def listen():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("🎤 Listening...")
        recognizer.adjust_for_ambient_noise(source)
        audio = recognizer.listen(source)

    try:
        text = recognizer.recognize_google(audio)
        print("You said:", text)
        return text.upper()
    except sr.UnknownValueError:
        return ""


# Find acronym meaning
def find_meaning(text, acronym_dict):
    for word in text.split():
        if word in acronym_dict:
            return acronym_dict[word]
        
    return "Sorry, I could not find the meaning."


# Text-to-Speech (ONLY ONCE)
def speak(text):
    print("🔊 Speaking:", text)
    engine = pyttsx3.init()
    engine.setProperty('rate', 150)
    engine.say(text)
    engine.runAndWait()



def main():
    acronyms = load_acronyms("acronym.xlsx")
    speak("Hello, ask me the meaning of any acronym.")

    while True:
        text = listen()
        time.sleep(0.2)

        if text in ["EXIT", "QUIT", "STOP"]:
            speak("Goodbye! Have a nice day")
            break

        meaning = find_meaning(text, acronyms)
        speak(meaning)

main()
