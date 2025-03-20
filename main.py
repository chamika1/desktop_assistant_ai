import speech_recognition as sr
from commands import process_command
from utilities import speak

# Capture Voice Command
def listen():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        recognizer.adjust_for_ambient_noise(source)
        audio = recognizer.listen(source)
        try:
            command = recognizer.recognize_google(audio)
            print("You said:", command)
            return command.lower()
        except sr.UnknownValueError:
            return ""

if __name__ == "__main__":
    speak("Hello! How can I assist you?")
    while True:
        command = listen()
        if command:
            process_command(command)
