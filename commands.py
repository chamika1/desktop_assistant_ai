import webbrowser
import os
import time
import pyautogui
from utilities import call_groq_api, speak, clear_conversation
import numpy as np
import pygetwindow as gw
from PIL import ImageGrab
import pytesseract
import subprocess
import psutil
import screen_brightness_control as sbc
from datetime import datetime
import win32com.client
import winreg
import json
import random
from pathlib import Path
from collections import deque
from threading import Thread, Event
from music_player import play_song
import urllib.parse

# Initialize data storage
DATA_DIR = Path.home() / "voice_assistant_data"
DATA_DIR.mkdir(exist_ok=True)
PLAYLISTS_FILE = DATA_DIR / "playlists.json"
FAVORITES_FILE = DATA_DIR / "favorites.json"
HISTORY_FILE = DATA_DIR / "history.json"

# Music queue and playback state
music_queue = deque()
current_playlist = None
is_playing = False
playback_thread = None
stop_playback = Event()

# Command aliases and shortcuts
COMMAND_ALIASES = {
    "play": ["start", "listen to", "put on", "begin", "resume", "unpause"],
    "stop": ["pause", "halt", "end", "freeze", "stop playing", "stop music", "stop video"],
    "next": ["skip", "forward", "next song", "next video", "skip this", "play next", "go next"],
    "previous": ["back", "prev", "last song", "previous song", "go back", "play previous", "rewind"],
    "volume up": ["louder", "increase volume", "turn it up", "volume higher", "raise volume", "make it louder"],
    "volume down": ["quieter", "decrease volume", "turn it down", "volume lower", "lower volume", "make it quieter"],
    "add to favorites": ["like", "favorite", "save song", "bookmark", "save this", "add to liked"],
    "create playlist": ["new playlist", "make playlist", "start playlist", "create new playlist"],
    "add to queue": ["queue", "add next", "play next", "queue up", "add to playlist", "line up"],
    "what's playing": ["what song", "current song", "now playing", "what's this song", "what is playing", "show current"],
    "clear queue": ["empty queue", "stop queue", "clear playlist", "remove all", "empty playlist"],
    "open": ["launch", "start", "run", "execute", "show"],
    "close": ["exit", "quit", "terminate", "shut down", "end program"],
    "search": ["find", "look for", "locate", "search for", "show me"],
}

# App name aliases
APP_ALIASES = {
    "notepad": ["note", "notes", "text editor", "editor", "notepad++"],
    "calculator": ["calc", "calculator app", "math", "compute"],
    "settings": ["system settings", "windows settings", "preferences", "control panel"],
    "file explorer": ["explorer", "files", "file manager", "my computer", "this pc"],
    "task manager": ["processes", "tasks", "system monitor", "performance"],
    "browser": ["chrome", "edge", "firefox", "internet", "web browser"],
    "paint": ["mspaint", "drawing", "image editor", "microsoft paint"],
    "word": ["microsoft word", "word processor", "winword", "office word"],
    "excel": ["microsoft excel", "spreadsheet", "office excel"],
    "mail": ["email", "outlook", "mail app", "electronic mail"],
    "photos": ["gallery", "images", "pictures", "photo viewer"],
    "camera": ["webcam", "camera app", "video camera"],
    "store": ["microsoft store", "app store", "windows store", "store app"],
    "terminal": ["command prompt", "cmd", "powershell", "console", "shell"],
    "clock": ["time", "alarms", "timer", "stopwatch"],
    "calendar": ["schedule", "appointments", "events", "planner"],
    "maps": ["microsoft maps", "windows maps", "navigation", "directions"],
    "weather": ["forecast", "temperature", "weather app", "climate"],
    "spotify": ["spotify app", "spotify player", "music player"],
    "vlc": ["vlc player", "media player", "video player"],
    "discord": ["discord app", "discord chat", "discord voice"],
    "zoom": ["zoom app", "zoom meeting", "zoom calls"],
    "skype": ["skype app", "skype calls", "microsoft skype"],
}

def load_data(file_path):
    if file_path.exists():
        with open(file_path, 'r') as f:
            return json.load(f)
    return {}

def save_data(data, file_path):
    with open(file_path, 'w') as f:
        json.dump(data, f, indent=4)

# Load existing data
playlists = load_data(PLAYLISTS_FILE)
favorites = load_data(FAVORITES_FILE)
history = load_data(HISTORY_FILE)

def add_to_history(song):
    """Add song to playback history"""
    if "playback_history" not in history:
        history["playback_history"] = []
    history["playback_history"].insert(0, {"song": song, "timestamp": datetime.now().isoformat()})
    if len(history["playback_history"]) > 100:  # Keep last 100 songs
        history["playback_history"].pop()
    save_data(history, HISTORY_FILE)

def add_to_queue(songs, clear_existing=False):
    """Add songs to the playback queue"""
    global music_queue
    if clear_existing:
        music_queue.clear()
    if isinstance(songs, str):
        music_queue.append(songs)
    else:
        music_queue.extend(songs)
    
    if not is_playing:
        start_playlist_playback()

def start_playlist_playback():
    """Start playing songs from the queue"""
    global is_playing, playback_thread, stop_playback
    
    if is_playing:
        return
    
    stop_playback.clear()
    is_playing = True
    
    def playback_worker():
        while not stop_playback.is_set() and music_queue:
            song = music_queue.popleft()
            add_to_history(song)
            play_youtube_video(song, auto_next=True)
            time.sleep(5)  # Wait between songs
        
        global is_playing
        is_playing = False
    
    playback_thread = Thread(target=playback_worker, daemon=True)
    playback_thread.start()

def stop_playlist_playback():
    """Stop the current playlist playback"""
    global is_playing, stop_playback
    stop_playback.set()
    is_playing = False
    music_queue.clear()
    control_youtube("stop")

def manage_music_library(command):
    global playlists, favorites, current_playlist
    try:
        if "create playlist" in command:
            playlist_name = command.replace("create playlist", "").strip()
            if playlist_name:
                playlists[playlist_name] = []
                save_data(playlists, PLAYLISTS_FILE)
                speak(f"Created playlist {playlist_name}")
                
        elif "add to playlist" in command:
            parts = command.replace("add to playlist", "").strip().split(" ", 1)
            if len(parts) == 2:
                playlist_name, song = parts
                if playlist_name in playlists:
                    playlists[playlist_name].append(song)
                    save_data(playlists, PLAYLISTS_FILE)
                    speak(f"Added {song} to playlist {playlist_name}")
                else:
                    speak(f"Playlist {playlist_name} not found")
                    
        elif "play playlist" in command:
            playlist_name = command.replace("play playlist", "").strip()
            if playlist_name in playlists and playlists[playlist_name]:
                current_playlist = playlist_name
                speak(f"Playing playlist {playlist_name}")
                add_to_queue(playlists[playlist_name], clear_existing=True)
            else:
                speak(f"Playlist {playlist_name} is empty or not found")
                
        elif "shuffle playlist" in command:
            playlist_name = command.replace("shuffle playlist", "").strip()
            if playlist_name in playlists and playlists[playlist_name]:
                current_playlist = playlist_name
                songs = playlists[playlist_name].copy()
                random.shuffle(songs)
                speak(f"Shuffling playlist {playlist_name}")
                add_to_queue(songs, clear_existing=True)
                
        elif "add to favorites" in command:
            song = command.replace("add to favorites", "").strip()
            if song:
                if "favorites" not in playlists:
                    playlists["favorites"] = []
                if song not in playlists["favorites"]:
                    playlists["favorites"].append(song)
                    save_data(playlists, PLAYLISTS_FILE)
                    speak(f"Added {song} to favorites")
                    
        elif "play favorites" in command:
            if "favorites" in playlists and playlists["favorites"]:
                current_playlist = "favorites"
                speak("Playing your favorite songs")
                add_to_queue(playlists["favorites"], clear_existing=True)
            else:
                speak("Your favorites playlist is empty")
                
        elif "add to queue" in command:
            song = command.replace("add to queue", "").strip()
            if song:
                add_to_queue(song)
                speak(f"Added {song} to queue")
                
        elif "clear queue" in command:
            stop_playlist_playback()
            speak("Cleared the playback queue")
            
        elif "what's playing" in command or "current song" in command:
            if is_playing and history["playback_history"]:
                current = history["playback_history"][0]["song"]
                speak(f"Currently playing {current}")
            else:
                speak("Nothing is playing right now")
                
        elif "play history" in command:
            if history.get("playback_history"):
                songs = [entry["song"] for entry in history["playback_history"][:10]]  # Last 10 songs
                speak("Playing your recently played songs")
                add_to_queue(songs, clear_existing=True)
            else:
                speak("No playback history found")
                
    except Exception as e:
        print(f"Error in music library management: {str(e)}")
        speak("Sorry, I had trouble managing the music library")
    
    return None

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def control_youtube(action):
    try:
        # First, ensure focus is on the YouTube window
        youtube_window = None
        for window in gw.getAllWindows():
            if "youtube" in window.title.lower():
                window.activate()
                youtube_window = window
                time.sleep(0.5)  # Wait for window activation
                break
        
        if action == "stop" or action == "pause":
            # More reliable pause using 'k' key (YouTube's official shortcut)
            pyautogui.press('k')
            time.sleep(0.2)
            speak("Stopping playback")
        elif action == "play" or action == "resume":
            # Resume playback using 'k' key if video is paused
            pyautogui.press('k')
            time.sleep(0.2)
            speak("Resuming playback")
        elif action == "next":
            # Use 'l' key to seek forward then 'SHIFT+N' for next video
            pyautogui.press('l')  # Seek forward first
            time.sleep(0.2)
            pyautogui.hotkey('shift', 'n')
            time.sleep(0.5)
            speak("Playing next video")
        elif action == "previous":
            # Use 'SHIFT+P' for previous video
            pyautogui.hotkey('shift', 'p')
            time.sleep(0.5)
            speak("Playing previous video")
            
        # Move mouse away to prevent overlay
        if youtube_window:
            pyautogui.moveTo(youtube_window.left + 10, youtube_window.top + 10)
            
    except Exception as e:
        print(f"Error controlling playback: {str(e)}")
        speak("Sorry, I couldn't control the playback")

def search_edge(query):
    try:
        # Open Edge with the search query
        os.system(f'start msedge "https://www.bing.com/search?q={query}"')
        speak(f"Searching for {query} in Edge")
    except Exception as e:
        print(f"Error searching in Edge: {str(e)}")

def search_shopping(query):
    try:
        # First try AliExpress
        os.system(f'start msedge "https://www.aliexpress.com/wholesale?SearchText={query}&SortType=price_asc"')
        speak(f"Searching for {query} on AliExpress, sorted by lowest price")
        time.sleep(2)
        
        # Then open eBay in new tab
        os.system(f'start msedge "https://www.ebay.com/sch/i.html?_nkw={query}&_sop=15"')
        speak("Also checking eBay prices")
        
    except Exception as e:
        print(f"Error searching for shopping: {str(e)}")

def play_on_streaming(service, title=None):
    try:
        if service == "netflix":
            if title:
                # Netflix search URL format
                os.system(f'start msedge "https://www.netflix.com/search?q={title}"')
                speak(f"Searching for {title} on Netflix")
                # Wait for page to load and click first result
                time.sleep(5)  # Netflix needs more time to load
                try:
                    # Click on the first search result
                    pyautogui.click(x=400, y=400)  # Adjust coordinates for Netflix
                    time.sleep(2)
                    # Click play button
                    pyautogui.click(x=500, y=500)  # Adjust coordinates for play button
                    speak(f"Playing {title}")
                except Exception as e:
                    print(f"Error clicking: {str(e)}")
            else:
                os.system('start msedge "https://www.netflix.com"')
                speak("Opening Netflix")

        elif service == "prime":
            if title:
                os.system(f'start msedge "https://www.primevideo.com/search/?k={title}"')
                speak(f"Searching for {title} on Prime Video")
                time.sleep(5)
                try:
                    pyautogui.click(x=400, y=400)  # Adjust coordinates for Prime
                    time.sleep(2)
                    pyautogui.click(x=500, y=500)  # Play button
                    speak(f"Playing {title}")
                except Exception as e:
                    print(f"Error clicking: {str(e)}")
            else:
                os.system('start msedge "https://www.primevideo.com"')
                speak("Opening Prime Video")

        elif service == "disney":
            if title:
                os.system(f'start msedge "https://www.disneyplus.com/search?q={title}"')
                speak(f"Searching for {title} on Disney Plus")
                time.sleep(5)
                try:
                    pyautogui.click(x=400, y=400)  # Adjust coordinates for Disney+
                    time.sleep(2)
                    pyautogui.click(x=500, y=500)  # Play button
                    speak(f"Playing {title}")
                except Exception as e:
                    print(f"Error clicking: {str(e)}")
            else:
                os.system('start msedge "https://www.disneyplus.com"')
                speak("Opening Disney Plus")

    except Exception as e:
        print(f"Error with streaming service: {str(e)}")

def analyze_screen():
    """Analyze screen content with proper error handling"""
    try:
        # Add a small delay to ensure window focus
        time.sleep(0.5)
        
        # Capture the screen
        screenshot = ImageGrab.grab()
        screenshot_np = np.array(screenshot)
        
        # Get list of all open windows
        windows = gw.getAllWindows()
        active_window = gw.getActiveWindow()
        
        # Prepare screen analysis report
        report = "I can see:\n"
        
        # Add active window info
        if active_window:
            report += f"Active window: {active_window.title}\n"
        
        # Add open windows info
        report += "\nOpen applications:\n"
        for window in windows:
            if window.visible:
                report += f"- {window.title}\n"
        
        # Try OCR only if Tesseract is available
        try:
            screen_text = pytesseract.image_to_string(screenshot_np)
            if screen_text.strip():
                report += "\nVisible text on screen:\n"
                report += screen_text
        except Exception as e:
            print(f"OCR error: {str(e)}")
            report += "\nCouldn't read text from screen"
        
        return report
        
    except Exception as e:
        print(f"Screen analysis error: {str(e)}")
        return "Sorry, I'm having trouble accessing the screen. Please check if I have screen capture permissions."

def control_system(command):
    try:
        if "volume" in command:
            if "up" in command or "increase" in command:
                for _ in range(5):
                    pyautogui.press('volumeup')
                speak("There you go, volume's up")
            elif "down" in command or "decrease" in command:
                for _ in range(5):
                    pyautogui.press('volumedown')
                speak("Alright, turned that down for you")
            elif "mute" in command:
                pyautogui.press('volumemute')
                speak("Muted!")
        elif "brightness" in command:
            try:
                if "up" in command or "increase" in command:
                    current = sbc.get_brightness()[0]
                    sbc.set_brightness(min(current + 20, 100))
                    speak("Brightness increased")
                elif "down" in command or "decrease" in command:
                    current = sbc.get_brightness()[0]
                    sbc.set_brightness(max(current - 20, 0))
                    speak("Brightness decreased")
                else:
                    level = int(''.join(filter(str.isdigit, command)))
                    sbc.set_brightness(level)
                    speak(f"Brightness set to {level}")
            except Exception as e:
                print(f"Brightness control error: {str(e)}")
        elif "wifi" in command:
            if "off" in command:
                os.system("netsh interface set interface 'Wi-Fi' admin=disable")
                speak("WiFi turned off")
            else:
                os.system("netsh interface set interface 'Wi-Fi' admin=enable")
                speak("WiFi turned on")
        elif "bluetooth" in command:
            if "off" in command:
                os.system("powershell Set-BluetoothStatus -Status Off")
                speak("Bluetooth turned off")
            else:
                os.system("powershell Set-BluetoothStatus -Status On")
                speak("Bluetooth turned on")
        elif "shutdown" in command:
            speak("Okay, I'll shut down the computer in 10 seconds. Let me know if you want to cancel")
            os.system("shutdown /s /t 10")
        elif "restart" in command:
            speak("Restarting computer in 10 seconds")
            os.system("shutdown /r /t 10")
        elif "sleep" in command:
            speak("Putting computer to sleep")
            os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
    except Exception as e:
        print(f"Error in system control: {str(e)}")

def manage_files(command):
    try:
        if "search file" in command:
            query = command.replace("search file", "").strip()
            os.system(f'explorer /select,"{query}"')
        elif "create folder" in command:
            folder_name = command.replace("create folder", "").strip()
            os.makedirs(folder_name)
        elif "delete" in command:
            path = command.split("delete")[-1].strip()
            if os.path.isfile(path):
                os.remove(path)
            elif os.path.isdir(path):
                os.rmdir(path)
        speak(f"File operation {command} completed")
    except Exception as e:
        print(f"Error in file management: {str(e)}")

def manage_calendar(command):
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        if "create event" in command:
            appt = outlook.CreateItem(1)  # 1 = Appointment
            # Parse command for details
            details = command.replace("create event", "").strip()
            appt.Subject = details
            appt.Save()
        elif "check appointments" in command:
            calendar = outlook.GetNamespace("MAPI").GetDefaultFolder(9)  # 9 = Calendar
            items = calendar.Items
            items.Sort("[Start]")
            today = datetime.now().date()
            appointments = [item for item in items if item.Start.date() == today]
            for appt in appointments:
                speak(f"You have {appt.Subject} at {appt.Start.strftime('%H:%M')}")
    except Exception as e:
        print(f"Error in calendar management: {str(e)}")

def get_app_path(app_name):
    # System apps and settings dictionary
    system_apps = {
        "settings": "ms-settings:",
        "control panel": "control",
        "calculator": "calc",
        "notepad": "notepad",
        "paint": "mspaint",
        "task manager": "taskmgr",
        "system information": "msinfo32",
        "disk cleanup": "cleanmgr",
        "character map": "charmap",
        "sound settings": "mmsys.cpl",
        "device manager": "devmgmt.msc",
        "disk management": "diskmgmt.msc",
        "services": "services.msc",
        "registry editor": "regedit",
        "group policy": "gpedit.msc",
        "remote desktop": "mstsc",
        "performance monitor": "perfmon",
        "resource monitor": "resmon",
        "system properties": "sysdm.cpl",
        "network connections": "ncpa.cpl",
        "power settings": "powercfg.cpl",
        "date and time": "timedate.cpl",
        "region settings": "intl.cpl",
        "user accounts": "netplwiz",
        "windows security": "windowsdefender:",
        "bluetooth settings": "ms-settings:bluetooth",
        "wifi settings": "ms-settings:network-wifi",
        "display settings": "ms-settings:display",
        "personalization": "ms-settings:personalization",
        "apps settings": "ms-settings:appsfeatures",
        "update settings": "ms-settings:windowsupdate",
        # Add common variations
        "calc": "calc",
        "calculator": "calc",
        "note": "notepad",
        "notes": "notepad",
        "wordpad": "write",
        "cmd": "cmd.exe",
        "command prompt": "cmd.exe",
        "terminal": "cmd.exe",
        "powershell": "powershell.exe",
        "photos": "ms-photos:",
        "camera": "microsoft.windows.camera:",
        "maps": "ms-windows-store:PDP?PFN=Microsoft.WindowsMaps",
        "clock": "ms-clock:",
        "alarms": "ms-clock:",
        "mail": "outlookmail:",
        "email": "outlookmail:",
        "store": "ms-windows-store:",
        "microsoft store": "ms-windows-store:"
    }

    # Convert input to lowercase for case-insensitive matching
    app_name_lower = app_name.lower()
    
    # Check if it's a system app (case-insensitive)
    for key, value in system_apps.items():
        if app_name_lower == key.lower():
            return value
    
    # Check common paths (case-sensitive for paths)
    common_paths = {
        "telegram": [
            r"%USERPROFILE%\AppData\Roaming\Telegram Desktop\Telegram.exe",
            r"%LOCALAPPDATA%\Telegram Desktop\Telegram.exe",
            r"C:\Program Files\Telegram Desktop\Telegram.exe",
            r"C:\Program Files (x86)\Telegram Desktop\Telegram.exe",
        ],
        "chrome": [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        ],
        "word": [
            r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
            r"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE",
        ],
        "excel": [
            r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
            r"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE",
        ]
    }
    
    # Check common paths (case-insensitive for app names)
    for key, paths in common_paths.items():
        if app_name_lower == key.lower():
            for path in paths:
                expanded_path = os.path.expandvars(path)
                if os.path.exists(expanded_path):
                    return f'"{expanded_path}"'
    
    # Try Windows Registry for other installed apps
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{}.exe".format(app_name)) as key:
            path = winreg.QueryValue(key, None)
            if os.path.exists(path):
                return f'"{path}"'
    except:
        pass
    
    return None

def play_youtube_video(search_term, retry_count=0, auto_next=False):
    """Enhanced YouTube video playback with retry mechanism"""
    try:
        # Use YouTube Music for better music playback if it's a song
        if any(word in search_term.lower() for word in ["song", "music", "audio"]):
            url = f"https://music.youtube.com/search?q={search_term}"
        else:
            # Add filters for better results: HD videos with good view count
            filters = "&sp=EgIQAXABUAFwAQ%253D%253D"  # HD + Most viewed
            url = f"https://www.youtube.com/results?search_query={search_term}{filters}"
        
        webbrowser.open(url)
        speak(f"Searching for {search_term}")
        
        # Wait for page to load with exponential backoff
        wait_time = 3 * (1.5 ** retry_count)  # Increases wait time with each retry
        time.sleep(wait_time)
        
        try:
            # Try multiple positions for more reliable clicking
            click_positions = [
                (400, 300),  # Default first video
                (400, 400),  # Second video
                (400, 200)   # Header video
            ]
            
            success = False
            for pos_x, pos_y in click_positions:
                try:
                    pyautogui.moveTo(pos_x, pos_y, duration=0.5)
                    # Check if cursor is over a video (look for hover effects)
                    time.sleep(0.5)
                    pyautogui.click()
                    success = True
                    break
                except:
                    continue
            
            if not success and retry_count < 2:
                # Retry with increased wait time
                return play_youtube_video(search_term, retry_count + 1, auto_next)
            elif success:
                speak(f"Playing {search_term}")
                if auto_next:
                    time.sleep(5)  # Wait between songs
                    return play_youtube_video(search_term, 0, auto_next)
            else:
                speak("Please select the video you want to play")
                
        except Exception as e:
            print(f"Error selecting video: {str(e)}")
            if retry_count < 2:
                return play_youtube_video(search_term, retry_count + 1, auto_next)
            else:
                speak("I found some videos but couldn't select automatically. Please click the one you want.")
    except Exception as e:
        print(f"Error in YouTube playback: {str(e)}")
        speak("Sorry, I had trouble playing that. Please try again.")

def normalize_command(command):
    """Convert various command phrasings to standard format"""
    command = command.lower().strip()
    
    # Remove common filler phrases
    filler_phrases = [
        "i said",
        "i want",
        "can you",
        "please",
        "hey",
        "you",
        "for me",
        "the song"
    ]
    
    for phrase in filler_phrases:
        command = command.replace(phrase, "").strip()
    
    # Check for command aliases
    for standard, alias_list in COMMAND_ALIASES.items():
        for alias in alias_list:
            if alias in command:
                command = command.replace(alias, standard)
                break
    
    # Smart command processing for music
    if any(word in command for word in ["play", "song", "music"]):
        # Extract just the song name
        for prefix in ["play song", "play", "song"]:
            if command.startswith(prefix):
                command = command.replace(prefix, "", 1).strip()
                command = f"song {command}"  # Normalize to standard format
                break
    
    # Check for app aliases if it's an open/close command
    if "open" in command or "close" in command:
        for app, aliases in APP_ALIASES.items():
            for alias in aliases:
                if alias in command:
                    # Replace the alias with the standard app name
                    command = command.replace(alias, app)
                    break
    
    # Smart command processing
    if "youtube" in command:
        command = command.replace("youtube", "").strip()
    
    # Handle number words
    number_words = {
        "one": "1", "two": "2", "three": "3", "four": "4", "five": "5",
        "six": "6", "seven": "7", "eight": "8", "nine": "9", "ten": "10"
    }
    for word, num in number_words.items():
        if word in command:
            command = command.replace(word, num)
    
    return command

def find_video_on_screen(song_name):
    """Find and click the correct video using screen analysis"""
    try:
        # Wait a bit longer for the page to fully load
        time.sleep(2)
        
        # Capture the screen
        screenshot = ImageGrab.grab()
        screenshot_np = np.array(screenshot)
        
        # Get text from screen with better configuration
        screen_text = pytesseract.image_to_string(
            screenshot_np,
            config='--psm 6'  # Assume uniform block of text
        )
        
        print("Screen text found:", screen_text)  # Debug print
        
        # Find all YouTube video titles
        video_entries = []
        y_offset = 250  # Starting Y position for first video
        y_increment = 100  # Typical spacing between videos
        
        lines = screen_text.split('\n')
        for i, line in enumerate(lines):
            line = line.strip()
            if len(line) > 10:  # Ignore very short lines
                y_pos = y_offset + (i * y_increment)
                video_entries.append((line, y_pos))
                print(f"Found video: {line} at y={y_pos}")
        
        # Find best match
        best_match = None
        best_y = None
        song_name_lower = song_name.lower()
        
        for title, y_pos in video_entries:
            if song_name_lower in title.lower():
                best_match = title
                best_y = y_pos
                print(f"Found matching video: {title} at y={y_pos}")
                break
        
        if best_match:
            # Click sequence for better reliability
            x_pos = 400  # X position for video titles
            
            # Single click sequence
            pyautogui.moveTo(x_pos, best_y, duration=0.5)
            time.sleep(0.5)
            pyautogui.click()
            
            # Don't press k - let YouTube autoplay handle it
            return True
            
        # Fallback to default position if no match found
        else:
            print("No exact match found, trying first video position")
            pyautogui.moveTo(400, 250, duration=0.5)
            time.sleep(0.5)
            pyautogui.click()
            return True
            
    except Exception as e:
        print(f"Error in screen analysis: {str(e)}")
        return False

def process_command(command):
    # Normalize command first
    command = normalize_command(command)
    
    # Handle song playback commands consistently
    if ("play song" in command) or ("song" in command):
        # Extract song name, handling both "play song X" and "song X" formats
        song_name = command.replace("play song", "").replace("song", "").strip()
        if song_name:
            # Use regular YouTube for more reliable playback
            search_query = urllib.parse.quote(song_name)
            url = f"https://www.youtube.com/results?search_query={search_query}"
            webbrowser.open(url)
            speak(f"Searching for {song_name}")
            
            # Wait for page to load
            time.sleep(3)
            
            # Try to find and click the correct video
            if find_video_on_screen(song_name):
                speak(f"Playing {song_name}")
            else:
                speak("I found some videos but couldn't select automatically. Please click the one you want.")
            return
        else:
            speak("Please specify a song to play")
            return
            
    # Quick commands (shortcuts)
    quick_commands = {
        "p": "pause",
        "n": "next",
        "b": "previous",
        "+": "volume up",
        "-": "volume down",
        "q": "add to queue",
        "f": "add to favorites"
    }
    
    if command in quick_commands:
        command = quick_commands[command]
    
    # Handle Windows Settings
    if any(phrase in command.lower() for phrase in ["open settings", "show settings", "launch settings"]):
        try:
            os.system("start ms-settings:")
            speak("Opening Windows Settings")
            return
        except Exception as e:
            print(f"Error opening settings: {str(e)}")
            speak("Sorry, I couldn't open Windows Settings")
            return

    # Clear conversation history
    if any(phrase in command.lower() for phrase in ["clear memory", "forget conversation", "start new conversation"]):
        clear_conversation()
        return

    # Screen analysis commands
    if any(phrase in command.lower() for phrase in [
        "what do you see", "what's on my screen", "what's open",
        "check my desktop", "analyze screen", "what applications are open"
    ]):
        try:
            screen_info = analyze_screen()
            print("Screen Analysis Result:", screen_info)  # Debug print
            speak(screen_info)
        except Exception as e:
            print(f"Error during screen analysis: {e}")
            speak("I'm having trouble analyzing the screen")
        return

    # Handle streaming services
    if any(word in command.lower() for word in ["watch", "movie", "netflix", "prime", "disney", "show", "series"]):
        # Extract title if present
        title = None
        if "watch" in command.lower():
            title = command.lower().replace("watch", "").strip()
            for service in ["on netflix", "on prime", "on disney"]:
                title = title.replace(service, "").strip()
        
        if "netflix" in command.lower():
            play_on_streaming("netflix", title)
            return
        elif "prime" in command.lower():
            play_on_streaming("prime", title)
            return
        elif "disney" in command.lower():
            play_on_streaming("disney", title)
            return
        else:
            # Default to Netflix if just "watch movie" is said
            play_on_streaming("netflix", title)
            return

    # Handle shopping commands
    if any(word in command.lower() for word in ["shop", "buy", "purchase", "find item"]):
        search_query = command.lower()
        search_query = search_query.replace("shop", "").replace("for", "")
        search_query = search_query.replace("buy", "").replace("purchase", "")
        search_query = search_query.replace("find", "").replace("item", "")
        search_query = search_query.strip()
        
        if search_query:
            search_shopping(search_query)
            return

    # Handle search commands
    if "search" in command.lower():
        search_query = command.lower().replace("search", "").replace("for", "").strip()
        if search_query:
            search_edge(search_query)
            return

    # Handle YouTube playback controls
    if any(word in command.lower() for word in ["stop", "pause", "play", "resume", "next", "previous"]):
        if "stop" in command.lower() or "pause" in command.lower():
            control_youtube("stop")
            return
        elif "play" in command.lower() or "resume" in command.lower():
            control_youtube("play")
            return
        elif "next" in command.lower():
            control_youtube("next")
            return
        elif "previous" in command.lower():
            control_youtube("previous")
            return

    # Handle YouTube and music commands
    if any(word in command.lower() for word in ["youtube", "play", "song", "music"]):
        # If just "open youtube"
        if command.lower() == "open youtube":
            webbrowser.open("https://www.youtube.com")
            return
        
        # Check if it's a playlist command first
        playlist_song = manage_music_library(command)
        if playlist_song:
            play_youtube_video(playlist_song)
            return
        
        # For play commands
        search_term = command.lower()
        search_term = search_term.replace("play", "")
        search_term = search_term.replace("on youtube", "")
        search_term = search_term.replace("music", "")
        search_term = search_term.replace("song", "")
        search_term = search_term.strip()
        
        if search_term:
            play_youtube_video(search_term)
        return

    # Handle application opening commands
    if "open" in command.lower():
        app_name = command.lower().replace("open", "").strip()
        try:
            if app_name in ["file manager", "file explorer", "explorer"]:
                os.system("explorer")
                speak("Here's your file explorer")
            else:
                app_path = get_app_path(app_name)
                if app_path:
                    if app_path.startswith('"'):
                        # For executable paths
                        subprocess.Popen(app_path)
                    elif app_path.startswith("ms-settings:"):
                        # For Windows Settings URLs
                        os.system(f"start {app_path}")
                    else:
                        # For system commands
                        os.system(app_path)
                    speak(f"Opening {app_name}")
                else:
                    speak(f"I couldn't find {app_name}. Mind checking if it's installed?")
            return
        except Exception as e:
            print(f"Error opening application: {str(e)}")
            speak(f"Sorry, I couldn't open {app_name}")

    # Handle application closing commands
    if "close" in command.lower():
        app_name = command.lower().replace("close", "").strip()
        try:
            if app_name in ["chrome", "edge", "firefox", "browser"]:
                os.system("taskkill /im msedge.exe /f")
                os.system("taskkill /im chrome.exe /f")
                speak(f"Closing {app_name}")
            elif app_name == "telegram":
                os.system("taskkill /im telegram.exe /f")
                speak("Closing telegram")
            elif app_name in ["word", "microsoft word"]:
                os.system("taskkill /im winword.exe /f")
                speak("Closing Microsoft Word")
            elif app_name in ["excel", "microsoft excel"]:
                os.system("taskkill /im excel.exe /f")
                speak("Closing Microsoft Excel")
            elif app_name == "all":
                # Close common applications
                os.system("taskkill /im msedge.exe /f")
                os.system("taskkill /im chrome.exe /f")
                os.system("taskkill /im telegram.exe /f")
                os.system("taskkill /im winword.exe /f")
                os.system("taskkill /im excel.exe /f")
                speak("Closing all applications")
            else:
                os.system(f"taskkill /im {app_name}.exe /f")
                speak(f"Attempting to close {app_name}")
            return
        except Exception as e:
            print(f"Error closing application: {str(e)}")

    # System Controls
    if any(word in command for word in ["volume", "brightness", "wifi", "bluetooth", "shutdown", "restart", "sleep"]):
        control_system(command)
        return

    # File Management
    if any(word in command for word in ["search file", "create folder", "delete file", "move file"]):
        manage_files(command)
        return

    # Calendar and Tasks
    if any(word in command for word in ["create event", "set reminder", "check appointments"]):
        manage_calendar(command)
        return

    # Music Library
    if any(word in command for word in ["create playlist", "add to playlist", "play playlist", "shuffle playlist", "add to favorites", "play favorites", "add to queue", "clear queue", "what's playing", "play history"]):
        manage_music_library(command)
        return

    # For all other commands, use AI response
    response = call_groq_api(command)
    speak(response)
