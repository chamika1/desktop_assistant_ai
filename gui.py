import tkinter as tk
from tkinter import ttk
import threading
from main import listen
from commands import process_command
from utilities import speak
import sv_ttk  # For modern theme

class VoiceAssistantGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Voice Assistant")
        self.root.geometry("800x600")
        
        # Apply modern theme
        sv_ttk.set_theme("dark")
        
        # Main frame
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Status frame
        self.status_frame = ttk.Frame(self.main_frame)
        self.status_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Status indicator
        self.status_label = ttk.Label(self.status_frame, text="🎤 Ready", font=('Segoe UI', 12))
        self.status_label.pack(side=tk.LEFT)
        
        # Command history
        self.history_frame = ttk.LabelFrame(self.main_frame, text="Conversation History", padding="10")
        self.history_frame.pack(fill=tk.BOTH, expand=True)
        
        # Text widget for history
        self.history_text = tk.Text(self.history_frame, wrap=tk.WORD, font=('Segoe UI', 10))
        self.history_text.pack(fill=tk.BOTH, expand=True)
        
        # Scrollbar for history
        scrollbar = ttk.Scrollbar(self.history_frame, command=self.history_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.history_text.configure(yscrollcommand=scrollbar.set)
        
        # Control buttons frame
        self.button_frame = ttk.Frame(self.main_frame)
        self.button_frame.pack(fill=tk.X, pady=(10, 0))
        
        # Start button
        self.start_button = ttk.Button(self.button_frame, text="Start Listening", command=self.start_listening)
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        # Stop button
        self.stop_button = ttk.Button(self.button_frame, text="Stop", command=self.stop_listening)
        self.stop_button.pack(side=tk.LEFT, padx=5)
        
        # Clear button
        self.clear_button = ttk.Button(self.button_frame, text="Clear History", command=self.clear_history)
        self.clear_button.pack(side=tk.RIGHT, padx=5)
        
        self.is_listening = False
        
    def start_listening(self):
        if not self.is_listening:
            self.is_listening = True
            self.status_label.config(text="🎤 Listening...")
            threading.Thread(target=self.listen_loop, daemon=True).start()
    
    def stop_listening(self):
        self.is_listening = False
        self.status_label.config(text="🎤 Ready")
    
    def listen_loop(self):
        while self.is_listening:
            command = listen()
            if command:
                self.add_to_history(f"You: {command}")
                response = process_command(command)
                if response:
                    self.add_to_history(f"Assistant: {response}")
    
    def add_to_history(self, text):
        self.history_text.insert(tk.END, f"{text}\n")
        self.history_text.see(tk.END)
    
    def clear_history(self):
        self.history_text.delete(1.0, tk.END)
    
    def run(self):
        speak("Hello! I'm ready to help.")
        self.root.mainloop()

if __name__ == "__main__":
    app = VoiceAssistantGUI()
    app.run() 