import speech_recognition as sr
from commands import process_command
from utilities import speak

def listen():
    """Enhanced listening function with better error handling and feedback"""
    recognizer = sr.Recognizer()
    
    # Adjust recognition parameters for better sensitivity
    recognizer.dynamic_energy_threshold = False  # Use fixed threshold instead
    recognizer.energy_threshold = 1000  # Lower threshold for easier voice detection
    recognizer.pause_threshold = 0.5    # Shorter pause for more natural speech
    
    with sr.Microphone() as source:
        print("Listening...")
        try:
            # Shorter ambient noise adjustment
            recognizer.adjust_for_ambient_noise(source, duration=0.5)
            
            # More generous listening duration
            audio = recognizer.listen(source, timeout=10, phrase_time_limit=10)
            
            try:
                # Try multiple recognition attempts
                command = recognizer.recognize_google(audio, language='en-US', show_all=False)
                print("You said:", command)
                return command.lower()
                
            except sr.UnknownValueError:
                # Only print message, don't speak to avoid interruption
                print("Could not understand audio, please try again")
            except sr.RequestError as e:
                print(f"Could not request results; {e}")
                speak("I'm having trouble connecting to speech recognition")
            
        except sr.WaitTimeoutError:
            print("Listening timed out")
        except Exception as e:
            print(f"Error in listening: {e}")
    
    return ""

def main():
    speak("Hello! How can I assist you?")
    
    while True:
        try:
            command = listen()
            if command:
                process_command(command)
                
        except KeyboardInterrupt:
            print("\nExiting...")
            break
        except Exception as e:
            print(f"Error in main loop: {e}")
            speak("Sorry, something went wrong. Please try again.")

if __name__ == "__main__":
    main()
