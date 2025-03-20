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

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def control_youtube(action):
    try:
        if action == "stop":
            # Press spacebar to pause/stop
            pyautogui.press('space')
            speak("Stopping playback")
        elif action == "next":
            # Press 'n' key for next video
            pyautogui.press('n')
            speak("Playing next video")
        elif action == "previous":
            # Press 'p' key for previous video
            pyautogui.press('p')
            speak("Playing previous video")
    except Exception as e:
        print(f"Error controlling playback: {str(e)}")

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
    try:
        # Capture the entire screen
        screenshot = ImageGrab.grab()
        screenshot_np = np.array(screenshot)
        
        # Get list of all open windows
        windows = gw.getAllWindows()
        active_window = gw.getActiveWindow()
        
        # Extract text from screen using OCR
        screen_text = pytesseract.image_to_string(screenshot_np)
        
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
        
        # Add any visible text from screen
        if screen_text:
            report += "\nVisible text on screen:\n"
            report += screen_text
        
        return report
    except Exception as e:
        print(f"Error analyzing screen: {str(e)}")
        return "I'm having trouble analyzing the screen."

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
    
    if app_name in common_paths:
        for path in common_paths[app_name]:
            expanded_path = os.path.expandvars(path)  # Expand environment variables
            if os.path.exists(expanded_path):
                return f'"{expanded_path}"'
    
    # If no path found, try Windows Registry for installed apps
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{}.exe".format(app_name)) as key:
            path = winreg.QueryValue(key, None)
            if os.path.exists(path):
                return f'"{path}"'
    except:
        pass
    
    app_path = None
    if not app_path:
        speak(f"Hey, I can't seem to find {app_name}. Is it installed?")
    return app_path

def process_command(command):
    # Clear conversation history
    if any(phrase in command.lower() for phrase in ["clear memory", "forget conversation", "start new conversation"]):
        clear_conversation()
        return

    # Add screen analysis commands
    if any(phrase in command.lower() for phrase in [
        "what do you see", "what's on my screen", "what's open",
        "check my desktop", "analyze screen", "what applications are open"
    ]):
        screen_info = analyze_screen()
        speak(screen_info)
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
    if any(word in command.lower() for word in ["stop", "pause", "next", "previous"]):
        if "stop" in command.lower() or "pause" in command.lower():
            control_youtube("stop")
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
        
        # For play commands
        search_term = command.lower()
        search_term = search_term.replace("play", "")
        search_term = search_term.replace("on youtube", "")
        search_term = search_term.replace("music", "")
        search_term = search_term.replace("song", "")
        search_term = search_term.strip()
        
        if search_term:
            webbrowser.open(f"https://www.youtube.com/results?search_query={search_term}")
            speak(f"Searching for {search_term}")
            
            # Wait longer for page to load and add safety checks
            time.sleep(5)  # Increased wait time
            try:
                # Move mouse slowly to avoid triggering failsafe
                pyautogui.moveTo(400, 300, duration=1)
                pyautogui.click()
                
                # Wait for video page to load
                time.sleep(2)
                
                # Click play button if needed
                pyautogui.press('k')  # YouTube's play/pause shortcut
                speak(f"Playing {search_term} on YouTube")
            except Exception as e:
                print(f"Error playing video: {str(e)}")
                speak("I found the video but couldn't play it automatically. Please click it manually.")
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
                    subprocess.Popen(app_path)
                    speak(f"There you go, opened {app_name} for you")
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

    # For all other commands, use AI response
    response = call_groq_api(command)
    speak(response)
