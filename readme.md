# Voice Assistant

A powerful voice assistant that can help you with various tasks including music playback, web searches, shopping, and more.

## Features

### 1. Voice Commands
- The assistant listens for your voice commands
- Speaks responses back to you
- Maintains conversation memory for contextual interactions

### 2. Music and YouTube Control
#### a) Music and YouTube:
- **Play Music:**
  - "play [song name]"
  - "play me a song"
  - "play random music"
  - "open youtube"
- **Playback Control:**
  - "stop" or "pause" - Stops current playback
  - "next" - Plays next video
  - "previous" - Plays previous video

#### b) Streaming Services:
- **Watch Movies/Shows:**
  - "watch [movie/show name]" (searches on Netflix)
  - "watch [title] on netflix"
  - "watch [title] on prime"
  - "watch [title] on disney"
  - "open netflix/prime/disney"

### 3. Application Control
- Open applications:
  - "open chrome"
  - "open telegram"
  - "open word"
  - "open excel"
  - "open file explorer"
  - "open [any app]"
- Close applications:
  - "close chrome" (or "close browser")
  - "close telegram"
  - "close word"
  - "close excel"
  - "close [app name]"
  - "close all" (closes common applications)

### 4. Web Search
- Search the web using Microsoft Edge:
  - "search [topic]"
  - "search for [topic]"

### 5. Shopping Assistant
- Find best prices across online stores:
  - "shop for [item]"
  - "find [item]"
  - "buy [item]"
- Automatically searches:
  - AliExpress (sorted by lowest price)
  - eBay (sorted by lowest price)

### 6. Conversation Memory
- Maintains context of conversations
- Commands:
  - "clear memory"
  - "forget conversation"
  - "start new conversation"

## Setup

### 1. Install Required Packages:
```bash
pip install speech_recognition
pip install pyttsx3
pip install pyautogui
pip install groq
pip install pipwin
pipwin install pyaudio
pip install opencv-python
pip install numpy
pip install Pillow
pip install pygetwindow
pip install pytesseract
```

### 2. Install Tesseract OCR:
   - Windows: Download and install from https://github.com/UB-Mannheim/tesseract/wiki
   - Linux: `sudo apt-get install tesseract-ocr`
   - Mac: `brew install tesseract`

[rest of README content...]

### 1. Screen Analysis
- View and analyze desktop content:
  - "what do you see"
  - "what's on my screen"
  - "what applications are open"
  - "check my desktop"
  - "analyze screen"
- The assistant can:
  - List open applications
  - Identify active window
  - Read visible text
  - Describe screen content

### 2. Configure API Key:
- Create a `config.py` file
- Add your Groq API key:
```python
GROQ_API_KEY = "your_groq_api_key_here"
```

## Usage Tips

1. **Starting the Assistant:**
   - Run `main.py`
   - Wait for "Hello! How can I assist you?"

2. **Speaking Commands:**
   - Speak clearly into your microphone
   - Wait for the "Listening..." prompt
   - The assistant will respond verbally

3. **Shopping:**
   - For best results, be specific with item descriptions
   - The assistant will search multiple stores
   - Results are automatically sorted by lowest price

4. **YouTube Control:**
   - Make sure the YouTube window is active for playback controls
   - Allow a few seconds for video loading
   - Use stop/pause before playing new songs

5. **Memory Features:**
   - The assistant remembers previous conversations
   - Clear memory if you want to start fresh
   - Limited to last 10 exchanges for optimal performance

## Troubleshooting

1. **If microphone isn't working:**
   - Check microphone permissions
   - Verify PyAudio installation
   - Try running as administrator

2. **If browser commands fail:**
   - Ensure Microsoft Edge is installed
   - Check internet connection
   - Verify no popup blockers are active

3. **If app opening fails:**
   - Verify the application is installed
   - Check if the app name is correct
   - Try running as administrator

## Note

Keep the following files in your project directory:
- main.py
- commands.py
- utilities.py
- config.py

## Contributing

Feel free to contribute to this project by:
1. Adding new features
2. Improving voice recognition
3. Adding more shopping sources
4. Enhancing the AI responses

# Voice Assistant Commands List

[... rest of README content ...]

