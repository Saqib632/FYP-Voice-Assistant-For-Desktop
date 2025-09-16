import pygetwindow as gw
search_mode_enabled = False
recording_thread = None

def is_chrome_running():
    # Checks if Chrome is running
    for proc in os.popen('tasklist').readlines():
        if 'chrome.exe' in proc:
            return True
    return False

import os
import time
import wmi
import cv2
import numpy as np
import requests
import pyttsx3
import pyautogui
import subprocess
import webbrowser
import speech_recognition as sr
import threading
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager

from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL

API_KEY = "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"
API_URL = "https://openrouter.ai/api/v1/chat/completions"

engine = pyttsx3.init()

def speak(text):
    print(f"Assistant: {text}")
    engine.say(text)
    engine.runAndWait()

def normalize_command(text):
    if not text:
        return ""
    normalized = text.lower().strip()
    normalized = normalized.replace("wi-fi", "wifi")
    normalized = normalized.replace("wi fi", "wifi")
    return normalized

def get_voice_command(prompt=None):
    recognizer = sr.Recognizer()
    
    recognizer.energy_threshold = 300
    recognizer.dynamic_energy_threshold = True
    recognizer.pause_threshold = 0.8
    recognizer.phrase_threshold = 0.3
    recognizer.non_speaking_duration = 0.8
    
    try:
        with sr.Microphone() as source:
            recognizer.adjust_for_ambient_noise(source, duration=0.5)
            if prompt:
                speak(prompt)
            print("Listening...")
            audio = recognizer.listen(source, timeout=10, phrase_time_limit=5)

            # Try Google Speech Recognition first
            try:
                command = normalize_command(recognizer.recognize_google(audio))
                print(f"Command: {command}")
                return command

            except sr.RequestError:
                # No internet available → try offline recognition
                print("No internet connection. Switching to offline recognition...")
                speak("No internet connection. Using offline mode.")
                try:
                    command = normalize_command(recognizer.recognize_sphinx(audio))
                    print(f"Command (offline): {command}")
                    return command
                except sr.UnknownValueError:
                    speak("Sorry, could not understand you in offline mode. You may type manually.")
                    return None
                except Exception as e:
                    speak(f"Offline recognition failed: {e}. Switching to text input.")
                    return get_text_input()

            except sr.UnknownValueError:
                speak("Sorry, I did not understand. Please try again.")
                return None

    except sr.WaitTimeoutError:
        speak("Listening timeout. Please try again.")
        return None
    except OSError as e:
        if "No default input device" in str(e):
            speak("No microphone found. Switching to text input.")
        else:
            speak(f"Microphone error: {e}. Switching to text input.")
        return get_text_input()
    except Exception as e:
        speak(f"Unexpected error: {e}. Switching to text input.")
        return get_text_input()

def get_text_input():
    """Fallback text input when speech recognition fails"""
    try:
        print("\n" + "="*50)
        print("SPEECH RECOGNITION UNAVAILABLE")
        print("Switching to text input mode...")
        print("Type your commands or 'exit' to quit")
        print("="*50)
        
        command = input("\nEnter command: ").strip()
        if command:
            return normalize_command(command)
        return None
    except KeyboardInterrupt:
        return "exit"
    except EOFError:
        return "exit"

def ask_openrouter(question):
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json"
    }
    data = {
        "model": "mistralai/mistral-7b-instruct",
        "messages": [{"role": "user", "content": question}]
    }
    try:
        response = requests.post(API_URL, headers=headers, json=data)
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"]
    except Exception as e:
        return f"Failed to get a response: {e}"

def open_item(path):
    try:
        os.startfile(path)
        speak(f"Opening {path}")
    except Exception as e:
        speak(f"Unable to open: {e}")

def open_website(url):
    chrome_path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
    webbrowser.register('chrome', None, webbrowser.BackgroundBrowser(chrome_path))
    webbrowser.get('chrome').open(f"https://{url}")
    speak(f"Opening {url.replace('.com','')}")

def close_app(app_name):
    try:
        os.system(f"taskkill /f /im {app_name}")
        speak(f"Closing {app_name}")
    except Exception as e:
        speak(f"Could not close {app_name}: {e}")

def change_volume(action):
    try:
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        current = volume.GetMasterVolumeLevelScalar()
        is_muted = volume.GetMute()

        if action == "increase":
            volume.SetMasterVolumeLevelScalar(min(current + 0.1, 1.0), None)
            speak("Volume increased.")
        elif action == "decrease":
            volume.SetMasterVolumeLevelScalar(max(current - 0.1, 0.0), None)
            speak("Volume decreased.")
        elif action == "mute":
            volume.SetMute(1, None)
            speak("Volume muted.")
        elif action == "unmute":
            if is_muted:
                volume.SetMute(0, None)
                speak("Volume unmuted.")
            else:
                speak("Already unmuted.")
    except Exception as e:
        speak(f"Volume error: {e}")

def change_brightness(action):
    try:
        wmi_obj = wmi.WMI(namespace='wmi')
        monitors = wmi_obj.WmiMonitorBrightnessMethods()
        current = wmi_obj.WmiMonitorBrightness()[0].CurrentBrightness
        for monitor in monitors:
            if action == "increase":
                monitor.WmiSetBrightness(min(current + 10, 100), 0)
                speak("Brightness increased.")
            elif action == "decrease":
                monitor.WmiSetBrightness(max(current - 10, 0), 0)
                speak("Brightness decreased.")
    except Exception as e:
        speak(f"Brightness error: {e}")

# ========== WIFI CONTROL ==========

def get_wifi_interface_name():
    try:
        result = subprocess.run(
            ["netsh", "wlan", "show", "interfaces"],
            capture_output=True,
            text=True,
            check=True
        )
        for line in result.stdout.splitlines():
            if line.strip().lower().startswith("name"):
                interface = line.split(":", 1)[1].strip()
                return interface if interface else "Wi-Fi"
    except subprocess.CalledProcessError:
        speak("Unable to get Wi-Fi interface name.")
    except Exception as e:
        speak(f"Error finding Wi-Fi interface: {e}")
    return "Wi-Fi"

def set_wifi_enabled(enabled):
    try:
        interface_name = get_wifi_interface_name()
        state = "enabled" if enabled else "disabled"
        command = f'netsh interface set interface name="{interface_name}" admin={state}'
        result = subprocess.run(command, shell=True, capture_output=True, text=True)

        if result.returncode == 0:
            speak(f"Wi-Fi {'turned on' if enabled else 'turned off'}.")
        else:
            speak(f"Failed to {'enable' if enabled else 'disable'} Wi-Fi. Try running as administrator.")
    except Exception as e:
        speak(f"Wi-Fi control error: {e}")

# ========== GOOGLE SEARCH MODE ==========

def google_search_mode():
    speak("What would you like to search for?")
    while True:
        query = get_voice_command()
        if any(x in query for x in ["exit", "stop", "cancel"]):
            speak("Exiting search.")
            break
        elif query:
            chrome_path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
            webbrowser.register('chrome', None, webbrowser.BackgroundBrowser(chrome_path))
            webbrowser.get('chrome').open(f"https://www.google.com/search?q={query}")
            speak(f"Searching for {query}")

# ========== CHROME TABS CONTROL ==========

def manage_chrome_tabs(action):
    try:
        if not is_chrome_running():
            speak("Chrome browser is not open. Please open Chrome first to manage tabs.")
            return
        # Bring Chrome to the foreground
        chrome_windows = [w for w in gw.getWindowsWithTitle('Google Chrome') if w.visible]
        if chrome_windows:
            chrome_windows[0].activate()
            time.sleep(0.2)  # Give time for window to focus
        else:
            speak("Could not find a visible Chrome window to focus.")
            return
        if action == "new":
            pyautogui.hotkey('ctrl', 't')
        elif action == "next":
            pyautogui.hotkey('ctrl', 'tab')
        elif action == "previous":
            pyautogui.hotkey('ctrl', 'shift', 'tab')
        elif action == "close":
            pyautogui.hotkey('ctrl', 'w')
        speak(f"Tab {action}")
    except Exception as e:
        speak(f"Tab control failed: {e}")

# ========== FOLDER CONTROL ==========

def open_folder(folder_name):
    try:
        folder_paths = {
            "downloads": os.path.join(os.path.expanduser("~"), "Downloads"),
            "music": os.path.join(os.path.expanduser("~"), "Music"),
            "videos": os.path.join(os.path.expanduser("~"), "Videos"),
            "desktop": os.path.join(os.path.expanduser("~"), "Desktop"),
            "documents": os.path.join(os.path.expanduser("~"), "Documents")
        }
        path = folder_paths.get(folder_name.lower())
        if path and os.path.exists(path):
            os.startfile(path)
            time.sleep(1)
            pyautogui.hotkey('win', 'up')
            speak(f"Opening {folder_name} folder in full screen.")
        else:
            speak(f"{folder_name} folder not found.")
    except Exception as e:
        speak(f"Error opening {folder_name}: {e}")

def close_folder(folder_name):
    try:
        os.system('taskkill /f /im explorer.exe')
        subprocess.Popen("explorer")
        speak(f"Closing {folder_name} folder.")
    except Exception as e:
        speak(f"Error closing folder: {e}")

# ========== WHATSAPP FEATURE ==========

def launch_whatsapp_web():
    speak("Opening WhatsApp Web. Please scan the QR code.")
    driver = webdriver.Chrome(ChromeDriverManager().install())
    driver.get("https://web.whatsapp.com")
    time.sleep(20)
    return driver

def send_whatsapp_message(driver, contact_name, message_text):
    try:
        search_box = driver.find_element(By.XPATH, "//div[@contenteditable='true'][@data-tab='3']")
        search_box.click()
        search_box.send_keys(contact_name)
        time.sleep(3)
        search_box.send_keys(Keys.ENTER)
        time.sleep(2)
        message_box = driver.find_element(By.XPATH, "//div[@contenteditable='true'][@data-tab='10']")
        message_box.click()
        message_box.send_keys(message_text)
        message_box.send_keys(Keys.ENTER)
        speak(f"Message sent to {contact_name}")
    except Exception as e:
        speak(f"Failed to send message to {contact_name}: {e}")

def start_whatsapp_chat():
    driver = launch_whatsapp_web()
    while True:
        contact = get_voice_command("Who do you want to message?")
        if contact in ["exit", "stop"]:
            break
        message = get_voice_command("What message do you want to send?")
        confirm = get_voice_command("Say 'send' to send the message")
        if "send" in confirm:
            send_whatsapp_message(driver, contact, message)
        else:
            speak("Cancelled sending.")

# ========== SCREEN RECORDING ==========

recording = False

def record_screen():
    global recording
    # Create a unique filename with a timestamp
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"Recording_{timestamp}.avi"
    
    # Get screen size
    screen_size = pyautogui.size()
    
    # Define the codec and create a VideoWriter object
    fourcc = cv2.VideoWriter_fourcc(*"XVID")
    out = cv2.VideoWriter(filename, fourcc, 20.0, screen_size)
    
    print(f"Recording started. Saving to {filename}")
    
    while recording:
        # Capture a screenshot
        img = pyautogui.screenshot()
        # Convert the screenshot to a numpy array
        frame = np.array(img)
        # Convert it from BGR(OpenCV default) to RGB
        frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        # Write the frame to the video file
        out.write(frame)

    # Release the VideoWriter and clean up
    out.release()
    print("Recording stopped and saved.")

def start_screen_recording():
    global recording, recording_thread
    if not recording:
        recording = True
        recording_thread = threading.Thread(target=record_screen)
        recording_thread.start()
        speak("Screen recording has started.")
    else:
        speak("Screen recording is already in progress.")

def stop_screen_recording():
    global recording
    if recording:
        recording = False
        # The thread will stop on its own when `recording` is False
        speak("Screen recording stopped and saved.")
    else:
        speak("No screen recording is currently active.")

# ========== SCREENSHOT ==========

def take_screenshot():
    """Takes a screenshot and saves it with a unique filename."""
    try:
        # Create a unique filename with a timestamp
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"Screenshot_{timestamp}.png"
        
        # Take the screenshot
        pyautogui.screenshot(filename)
        speak(f"Screenshot taken and saved as {filename}")
    except Exception as e:
        speak(f"Sorry, I failed to take a screenshot. Error: {e}")

# ========== MAIN HANDLER ==========

def handle_command(command):
    command = normalize_command(command)
    global search_mode_enabled
    if "enable search mode" in command:
        if is_chrome_running():
            search_mode_enabled = True
            speak("Search mode enabled. Say your search queries.")
        else:
            speak("Chrome browser is not open. Please open Chrome first.")
    elif "disable search mode" in command:
        if search_mode_enabled:
            search_mode_enabled = False
            speak("Search mode disabled.")
        else:
            speak("Search mode is already disabled.")
    elif search_mode_enabled:
        # Only process search queries in search mode
        if any(x in command for x in ["exit", "stop", "cancel"]):
            search_mode_enabled = False
            speak("Exiting search mode.")
        elif command:
            chrome_path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
            webbrowser.register('chrome', None, webbrowser.BackgroundBrowser(chrome_path))
            webbrowser.get('chrome').open(f"https://www.google.com/search?q={command}")
            speak(f"Searching for {command}")
    # ...existing code...
        start_whatsapp_chat()
    elif "open youtube" in command:
        open_website("youtube.com")
    elif "open google" in command or command == "google":
        open_website("google.com")
        google_search_mode()
    elif "search" == command.strip():
        google_search_mode()
    elif "open facebook" in command:
        open_website("facebook.com")
    elif "open github" in command:
        open_website("github.com")
    elif "open calculator" in command:
        subprocess.Popen(["calc.exe"])
        speak("Opening calculator.")
    elif "close calculator" in command:
        close_app("CalculatorApp.exe")
    elif "open notepad" in command:
        subprocess.Popen(["notepad.exe"])
    elif "close notepad" in command:
        close_app("notepad.exe")
    elif command.startswith("increase volume"):
        change_volume("increase")
    elif command.startswith("decrease volume"):
        change_volume("decrease")
    elif command.startswith("unmute"):
        change_volume("unmute")
    elif command.startswith("mute"):
        change_volume("mute")
    elif "increase brightness" in command:
        change_brightness("increase")
    elif "decrease brightness" in command:
        change_brightness("decrease")
    elif any(x in command for x in ["turn on wifi", "enable wifi", "wifi on", "turn on wi-fi", "enable wi-fi", "wi-fi on"]):
        set_wifi_enabled(True)
    elif any(x in command for x in ["turn off wifi", "disable wifi", "wifi off", "turn off wi-fi", "disable wi-fi", "wi-fi off", "off wifi"]):
        set_wifi_enabled(False)
    elif "new tab" in command:
        manage_chrome_tabs("new")
    elif "next tab" in command:
        manage_chrome_tabs("next")
    elif "previous tab" in command:
        manage_chrome_tabs("previous")
    elif "close tab" in command:
        manage_chrome_tabs("close")
    elif "search for" in command or "what is the weather" in command or "weather today" in command or command.strip() == "weather":
        if "search for" in command:
            query = command.replace("search for", "").strip()
            if not query:
                query = "weather"
        elif "what is the weather" in command:
            query = "weather"
        elif "weather today" in command:
            query = "weather today"
        elif command.strip() == "weather":
            query = "weather"
        else:
            query = command
        speak(f"Searching for {query}")
        chrome_path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
        webbrowser.register('chrome', None, webbrowser.BackgroundBrowser(chrome_path))
        webbrowser.get('chrome').open(f"https://www.google.com/search?q={query}")
    elif "what is" in command or "who is" in command or "define" in command:
        response = ask_openrouter(command)
        speak(response)
    elif command.startswith("open"):
        for folder in ["downloads", "music", "videos", "desktop", "documents"]:
            if folder in command:
                open_folder(folder)
                return
    elif command.startswith("close"):
        for folder in ["downloads", "music", "videos", "desktop", "documents"]:
            if folder in command:
                close_folder(folder)
                return
    elif "start screen recording" in command:
        start_screen_recording()
    elif "stop screen recording" in command:
        stop_screen_recording()
    elif "take a screenshot" in command or "capture screen" in command:
        take_screenshot()
    elif "exit" in command or "stop" in command:
        speak("Goodbye!")
        exit()
    else:
        speak("Sorry, I don't understand that command.")

def main():
    speak("Voice assistant is now active.")
    print("Voice assistant is now active.")
    print("Note: Turning off WiFi will switch to offline mode or text input.")
    print("Press Ctrl+C to exit.\n")
    
    consecutive_failures = 0
    while True:
        try:
            command = get_voice_command()
            if command is None:
                print("Speech recognition failed. Please try again.")
                time.sleep(1)
                continue
            elif command:
                handle_command(command)
            else:
                continue
        except KeyboardInterrupt:
            print("\nExiting voice assistant...")
            speak("Goodbye!")
            break
        except Exception as e:
            print(f"Unexpected error: {e}")
            speak("An error occurred. Please try again.")
            time.sleep(1)

if __name__ == "__main__":
    main()
