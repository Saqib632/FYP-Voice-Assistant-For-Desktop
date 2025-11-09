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
import queue
import pyautogui
import subprocess
import webbrowser
import speech_recognition as sr
import threading
from datetime import datetime
import winreg
import ctypes
from PIL import ImageGrab
try:
    import pytesseract
    HAS_PYTESSERACT = True
except Exception:
    HAS_PYTESSERACT = False

# WhatsApp Desktop integration (keyboard automation)
whatsapp_current_chat = None
last_message_text = None
awaiting_contact_search = False
awaiting_contact_selection = False
_whatsapp_search_query = None
whatsapp_suggestions = []
messaging_mode = None  # 'text' when in text-typing mode
messaging_thread = None

def _focus_whatsapp_window(timeout=8):
    """Bring WhatsApp window to the foreground. Return True if focused."""
    for _ in range(int(timeout/0.5)):
        wins = [w for w in gw.getWindowsWithTitle('WhatsApp') if w.visible]
        if wins:
            try:
                win = wins[0]
                win.activate()
                time.sleep(0.3)
                return True
            except Exception:
                pass
        time.sleep(0.5)
    return False

def open_whatsapp_app():
    """Open WhatsApp Desktop via Start menu if not already open, then focus it."""
    # Try to focus first
    if _focus_whatsapp_window():
        speak('WhatsApp is already open.')
        return True

    # Open Start, type WhatsApp, press Enter
    try:
        pyautogui.press('win')
        time.sleep(0.5)
        pyautogui.typewrite('WhatsApp', interval=0.05)
        time.sleep(0.3)
        pyautogui.press('enter')
        # wait for app to appear
        for _ in range(20):
            if _focus_whatsapp_window():
                speak('Opened WhatsApp.')
                return True
            time.sleep(0.5)
    except Exception as e:
        speak(f'Failed to open WhatsApp: {e}')
    speak('Could not open WhatsApp. Please open it manually and try again.')
    return False

def _search_contact_and_prepare(query):
    """Type the query into WhatsApp's search (Ctrl+K) and leave results visible.
    We do not parse UI; we leave the UI for the user and set flags for selection.
    """
    if not _focus_whatsapp_window():
        speak('WhatsApp window not found. Please open WhatsApp first.')
        return False
    try:
        pyautogui.hotkey('ctrl', 'k')
        time.sleep(0.3)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('backspace')
        time.sleep(0.1)
        pyautogui.typewrite(query, interval=0.03)
        time.sleep(0.6)
        return True
    except Exception as e:
        speak(f'Contact search failed: {e}')
        return False

def _open_top_search_result():
    """Assumes search results visible; pressing Enter opens the top result."""
    try:
        pyautogui.press('enter')
        time.sleep(0.6)
        return True
    except Exception:
        return False


def _open_search_result_by_index(index=1):
    """Open the nth search result (1-based) by sending Down arrows then Enter.
    This is more reliable than pressing Enter while the search box is focused.
    """
    # Try key navigation first (Down arrows then Enter)
    try:
        steps = max(1, index)
        for _ in range(steps):
            pyautogui.press('down')
            time.sleep(0.12)
        pyautogui.press('enter')
        time.sleep(0.6)
        return True
    except Exception:
        pass

    # Fallback: try OCR-based click helper (which itself falls back to coordinates)
    return _click_contact_by_name_or_index(name=None, index=index)

def _type_and_send_text_app(text):
    """Type into the message box of WhatsApp Desktop and press Enter."""
    if not _focus_whatsapp_window():
        speak('WhatsApp window not found. Please open WhatsApp first.')
        return False
    try:
        # Ensure message input is focused: press Tab a few times if needed
        # In most WhatsApp Desktop installs, after opening a chat the message box is focused
        pyautogui.typewrite(text, interval=0.02)
        pyautogui.press('enter')
        return True
    except Exception as e:
        speak(f'Failed to type message: {e}')
        return False

def _text_message_worker():
    """Background loop to listen and type messages while in text mode.
    This worker catches exceptions from the speech recognizer so it doesn't
    crash, and falls back to text input when voice fails.
    """
    global messaging_mode, last_message_text, messaging_thread
    stop_variants = (
        'stop text mode', 'exit text mode', 'disable text message', 'disable text',
        'disable text mode', 'turn off text message', 'stop text message'
    )
    speak('Text message mode started. Say "stop text mode" to finish.')
    try:
        while messaging_mode == 'text':
            try:
                cmd = get_voice_command()
            except Exception as e:
                # Catch any exceptions from the recognizer and switch to text input
                speak(f'Voice input error: {e}. Switching to text input for this message.', wait=False)
                cmd = get_text_input()

            if not cmd:
                continue
            # normalize for matching
            norm = cmd.lower()
            # stop if any stop variant appears
            if any(variant in norm for variant in stop_variants):
                messaging_mode = None
                speak('Exiting text message mode.', wait=False)
                break
            if 'forward message' in norm:
                # handled by global forward flow; ignore here to avoid inserting text
                handle_command('forward message')
                continue
            # don't allow the literal trigger/stop words to be typed
            safe = cmd
            for v in ('forward message',) + stop_variants:
                safe = safe.replace(v, '')
            safe = safe.strip()
            if safe:
                ok = _type_and_send_text_app(safe)
                if ok:
                    last_message_text = safe
    finally:
        # clear thread reference when worker exits
        try:
            messaging_thread = None
        except Exception:
            pass

def close_whatsapp_app():
    """Close WhatsApp Desktop gracefully (Alt+F4) and fall back to taskkill."""
    try:
        # Try graceful close
        focused = _focus_whatsapp_window(timeout=3)
        if focused:
            try:
                pyautogui.hotkey('alt', 'f4')
                time.sleep(0.6)
                # check if still present
                if not _focus_whatsapp_window(timeout=2):
                    speak('WhatsApp closed.')
                    return True
            except Exception:
                pass

        # Fallback: try common process names
        killed_any = False
        for proc_name in ("WhatsApp.exe", "WhatsAppDesktop.exe", "WhatsApp.exe"):
            try:
                rc = os.system(f'taskkill /IM "{proc_name}" /F >nul 2>&1')
                if rc == 0:
                    killed_any = True
            except Exception:
                pass

        if killed_any:
            speak('WhatsApp closed.')
            return True
        else:
            speak('Could not close WhatsApp programmatically. Please close it manually.')
            return False
    except Exception as e:
        speak(f'Error closing WhatsApp: {e}')
        return False

def _get_whatsapp_suggestions_via_ocr(max_suggestions=5):
    """Take a screenshot of the left pane of the WhatsApp window and OCR contact names.
    Returns a list of suggestion strings (may be empty). Requires pytesseract.
    """
    if not HAS_PYTESSERACT:
        return []
    wins = [w for w in gw.getWindowsWithTitle('WhatsApp') if w.visible]
    if not wins:
        return []
    win = wins[0]
    try:
        # Crop a reasonable left pane where contacts appear
        left = win.left + 10
        top = win.top + 80
        right = win.left + int(min(420, win.width * 0.5))
        bottom = win.top + int(min(win.height - 60, 600))
        img = ImageGrab.grab(bbox=(left, top, right, bottom))
        gray = img.convert('L')
        text = pytesseract.image_to_string(gray)
        # Split lines and clean
        lines = [l.strip() for l in text.splitlines() if l.strip()]
        # Heuristic: return up to max_suggestions unique lines
        seen = []
        for l in lines:
            if l not in seen:
                seen.append(l)
            if len(seen) >= max_suggestions:
                break
        return seen
    except Exception:
        return []


def _click_contact_by_name_or_index(name=None, index=1):
    """Use OCR to find a contact line by name (or phone digits) and click it.
    If OCR is not available or fails, fall back to approximate coordinate click.
    """
    # If pytesseract is not available, fallback to coordinate click
    if not HAS_PYTESSERACT:
        try:
            # reuse the coordinate click logic from fallback
            wins = [w for w in gw.getWindowsWithTitle('WhatsApp') if w.visible]
            if not wins:
                return False
            win = wins[0]
            left = win.left + 10
            top = win.top + 80
            x = left + 60
            y = top + 60 + (index-1) * 80
            pyautogui.moveTo(x, y, duration=0.12)
            pyautogui.click()
            time.sleep(0.6)
            return True
        except Exception:
            return False

    wins = [w for w in gw.getWindowsWithTitle('WhatsApp') if w.visible]
    if not wins:
        return False
    win = wins[0]
    try:
        left = win.left + 10
        top = win.top + 80
        right = win.left + int(min(420, win.width * 0.5))
        bottom = win.top + int(min(win.height - 60, 600))
        img = ImageGrab.grab(bbox=(left, top, right, bottom))
        gray = img.convert('L')
        # Use detailed OCR output to get bounding boxes
        data = pytesseract.image_to_data(gray, output_type=pytesseract.Output.DICT)
        n_boxes = len(data['level'])
        # Reconstruct lines by line_num
        lines = {}
        for i in range(n_boxes):
            text = (data['text'][i] or '').strip()
            if not text:
                continue
            ln = data['line_num'][i]
            x = int(data['left'][i])
            y = int(data['top'][i])
            w = int(data['width'][i])
            h = int(data['height'][i])
            if ln not in lines:
                lines[ln] = {'text': text, 'left': x, 'top': y, 'right': x + w, 'bottom': y + h}
            else:
                # append text and expand bbox
                lines[ln]['text'] += ' ' + text
                lines[ln]['left'] = min(lines[ln]['left'], x)
                lines[ln]['top'] = min(lines[ln]['top'], y)
                lines[ln]['right'] = max(lines[ln]['right'], x + w)
                lines[ln]['bottom'] = max(lines[ln]['bottom'], y + h)

        # Normalize search target
        target = name.lower().strip() if name else None
        target_digits = ''.join([c for c in target if c.isdigit()]) if target else None

        # Try to find a matching line by substring match
        for ln, info in lines.items():
            txt = info['text'].lower()
            if target and target in txt:
                # click center
                cx = left + info['left'] + (info['right'] - info['left']) // 2
                cy = top + info['top'] + (info['bottom'] - info['top']) // 2
                pyautogui.moveTo(cx, cy, duration=0.12)
                pyautogui.click()
                time.sleep(0.6)
                return True
            if target_digits and target_digits in ''.join([c for c in txt if c.isdigit()]):
                cx = left + info['left'] + (info['right'] - info['left']) // 2
                cy = top + info['top'] + (info['bottom'] - info['top']) // 2
                pyautogui.moveTo(cx, cy, duration=0.12)
                pyautogui.click()
                time.sleep(0.6)
                return True

        # If no name match, click by visual index (first visible line is index=1)
        sorted_lines = sorted(lines.items(), key=lambda x: x[0])
        if 1 <= index <= len(sorted_lines):
            info = sorted_lines[index-1][1]
            cx = left + info['left'] + (info['right'] - info['left']) // 2
            cy = top + info['top'] + (info['bottom'] - info['top']) // 2
            pyautogui.moveTo(cx, cy, duration=0.12)
            pyautogui.click()
            time.sleep(0.6)
            return True

        # fallback coordinate click if no lines found
        x = left + 60
        y = top + 60 + (index-1) * 80
        pyautogui.moveTo(x, y, duration=0.12)
        pyautogui.click()
        time.sleep(0.6)
        return True
    except Exception:
        return False

_number_words = {
    'one': 1, 'two': 2, 'three': 3, 'four': 4, 'five': 5,
    '1': 1, '2': 2, '3': 3, '4': 4, '5': 5
}

# include common ordinal and word variants so users can say "first", "second", "1st", etc.
_number_words.update({
    'first': 1, 'second': 2, 'third': 3, 'fourth': 4, 'fifth': 5,
    '1st': 1, '2nd': 2, '3rd': 3, '4th': 4, '5th': 5
})





from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL

API_KEY = "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"
API_URL = "https://openrouter.ai/api/v1/chat/completions"

# TTS queue and worker to serialize pyttsx3 calls and avoid concurrent run loop errors
_tts_queue = queue.Queue()


def _tts_worker():
    try:
        t_engine = pyttsx3.init()
    except Exception:
        t_engine = None
    while True:
        text, evt = _tts_queue.get()
        try:
            if t_engine is None:
                try:
                    t_engine = pyttsx3.init()
                except Exception:
                    t_engine = None
            if t_engine:
                t_engine.say(text)
                t_engine.runAndWait()
        except Exception:
            # If TTS fails, ignore and continue
            pass
        finally:
            if evt:
                try:
                    evt.set()
                except Exception:
                    pass
            _tts_queue.task_done()


# start the TTS worker thread
_tts_thread = threading.Thread(target=_tts_worker, daemon=True)
_tts_thread.start()


def speak(text, wait=True):
    """Speak text using a background TTS worker.
    If wait=True the call blocks until the utterance is finished.
    """
    try:
        print(f"Assistant: {text}")
        # Prevent deadlock: if speak is called from the TTS thread itself, do not wait
        current = threading.current_thread()
        if wait and '_tts_thread' in globals() and current is _tts_thread:
            wait = False
        evt = threading.Event() if wait else None
        _tts_queue.put((text, evt))
        if wait and evt:
            evt.wait()
    except Exception:
        # best-effort: fallback to printing only
        try:
            print(f"Assistant(FAIL): {text}")
        except Exception:
            pass

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


def close_chrome():
    """Close Google Chrome processes.

    Tries a graceful termination using psutil if available, waits briefly,
    and falls back to taskkill /F to ensure processes are terminated.
    """
    try:
        # Try graceful termination using psutil if installed
        try:
            import psutil
            chrome_procs = []
            for p in psutil.process_iter(['name']):
                name = p.info.get('name')
                if name and 'chrome' in name.lower():
                    chrome_procs.append(p)

            if chrome_procs:
                for p in chrome_procs:
                    try:
                        p.terminate()
                    except Exception:
                        pass

                gone, alive = psutil.wait_procs(chrome_procs, timeout=5)
                # kill any remaining
                for p in alive:
                    try:
                        p.kill()
                    except Exception:
                        pass

                speak("Closed Google Chrome.")
                return
        except Exception:
            # psutil not available or failed — fall back to taskkill
            pass

        # Fallback: force kill using taskkill
        os.system('taskkill /IM chrome.exe /F')
        speak("Closed Google Chrome.")
    except Exception as e:
        speak(f"Failed to close Chrome: {e}")

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

# ========== WINDOW THEME CONTROL ==========
def set_windows_theme(dark=True):
    """Set Windows system and apps theme to dark (True) or light (False).
    This writes to the current user's Personalize registry keys and broadcasts
    the setting change so the UI updates.
    """
    try:
        personalize_key = r"Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize"
        # 0 = dark, 1 = light
        value = 0 if dark else 1
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, personalize_key, 0, winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, "AppsUseLightTheme", 0, winreg.REG_DWORD, value)
            winreg.SetValueEx(key, "SystemUsesLightTheme", 0, winreg.REG_DWORD, value)

        # Notify the system about the change
        HWND_BROADCAST = 0xFFFF
        WM_SETTINGCHANGE = 0x001A
        SMTO_ABORTIFHUNG = 0x0002
        result = ctypes.c_ulong()
        ctypes.windll.user32.SendMessageTimeoutW(HWND_BROADCAST, WM_SETTINGCHANGE, 0,
                                                 ctypes.c_wchar_p("ImmersiveColorSet"), SMTO_ABORTIFHUNG, 5000,
                                                 ctypes.byref(result))
        return True
    except Exception as e:
        speak(f"Failed to change theme: {e}")
        return False


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
    global awaiting_contact_search, awaiting_contact_selection, _whatsapp_search_query, whatsapp_current_chat, messaging_mode, messaging_thread
    # Immediate overrides: if user asks to close WhatsApp at any point, do it first
    if any(x in command for x in [
        "close whatsapp", "close the whatsapp", "exit whatsapp", "quit whatsapp", "close the app",
        # common speech/recognition variants
        "close whatapp", "close what app", "close whats app", "close whatsapp", "close whatsapp app"
    ]):
        # clear any pending search/selection state to avoid typing into the search box
        awaiting_contact_search = False
        awaiting_contact_selection = False
        _whatsapp_search_query = None
        whatsapp_suggestions.clear() if 'whatsapp_suggestions' in globals() else None
        close_whatsapp_app()
        return
    # If awaiting a contact search query (after "open whatsapp") treat the spoken command as the query
    if awaiting_contact_search:
        q = command.strip()
        if q in ("cancel","exit","stop"):
            awaiting_contact_search = False
            speak('Cancelled contact search.')
            return
        ok = _search_contact_and_prepare(q)
        if ok:
            _whatsapp_search_query = q
            awaiting_contact_selection = True
            awaiting_contact_search = False
            # Try to OCR suggestions from WhatsApp window (best-effort)
            suggestions = _get_whatsapp_suggestions_via_ocr()
            if suggestions:
                # store and speak numbered list
                whatsapp_suggestions.clear()
                whatsapp_suggestions.extend(suggestions)
                speak('I found the following contacts:')
                for i, name in enumerate(whatsapp_suggestions, start=1):
                    speak(f'{i}. {name}')
                speak('Say the number of the contact to open it, or say the name exactly.')
            else:
                speak(f'I searched for {q}. I could not read suggestions automatically, please say the exact contact name to open it.')
        return

    if awaiting_contact_selection:
        # allow numbered selection if suggestions were found
        # Check for a spoken number word or digit
        tokens = command.lower().split()
        chosen_index = None
        for t in tokens:
            if t in _number_words:
                chosen_index = _number_words[t]
                break
        if chosen_index is not None:
            # If we have OCR-derived suggestions, prefer opening the selected suggestion
            if whatsapp_suggestions:
                if 1 <= chosen_index <= len(whatsapp_suggestions):
                    chosen_name = whatsapp_suggestions[chosen_index-1]
                    ok = _search_contact_and_prepare(chosen_name)
                    if ok and _open_search_result_by_index(1):
                        whatsapp_current_chat = chosen_name
                        awaiting_contact_selection = False
                        whatsapp_suggestions.clear() # Clear suggestions after selection
                        speak(f'Opening chat with {chosen_name}.')
                        return
                    else:
                        speak(f'Sorry, I could not open the chat for {chosen_name}.')
                        return
            else:
                # No OCR suggestions available, but user said "one/first/2/second" — perform key presses
                try:
                    if chosen_index <= 1:
                        # open top result
                        if _open_top_search_result():
                            awaiting_contact_selection = False
                            speak('Opening the first search result.')
                            return
                        else:
                            speak('Could not open the first result.')
                            return
                    else:
                        # No OCR suggestions available; open by index using arrow keys
                        try:
                            if _open_search_result_by_index(chosen_index):
                                awaiting_contact_selection = False
                                speak(f'Opening result number {chosen_index}.')
                                return
                            else:
                                speak('Could not open that result.')
                                return
                        except Exception as e:
                            speak(f'Failed to select result: {e}')
                            return
                except Exception as e:
                    speak(f'Failed to select result: {e}')
                    return

        # otherwise assume user spoke the exact contact name; type it and press Enter
        q = command.strip()
        if q in ("cancel","exit","stop"):
            awaiting_contact_selection = False
            whatsapp_suggestions.clear()
            speak('Cancelled contact selection.')
            return
        ok = _search_contact_and_prepare(q)
        if ok:
            # try robust open by index (uses keyboard nav + OCR click fallback)
            if _open_search_result_by_index(1):
                whatsapp_current_chat = q
                awaiting_contact_selection = False
                whatsapp_suggestions.clear()
                speak(f'Opening chat with {q}.')
            else:
                speak('Could not open that contact.')
        return
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
    # ...existing code... (WhatsApp feature removed)
    elif "open youtube" in command:
        open_website("youtube.com")
    elif "open whatsapp" in command or command.strip() == "whatsapp":
        if open_whatsapp_app():
            # ask for contact next
            awaiting_contact_search = True
            speak('Who do you want to message? Say the contact name.')
        return
    elif "close whatsapp" in command or "exit whatsapp" in command or "quit whatsapp" in command:
        close_whatsapp_app()
        return
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
    elif "close chrome" in command or "close google chrome" in command:
        close_chrome()
    if any(x in command for x in ["disable text message", "disable text", "disable text mode", "turn off text message", "stop text message"]):
        # stop text mode if active
        if messaging_mode == 'text':
            messaging_mode = None
            # give worker a moment to exit
            if messaging_thread and messaging_thread.is_alive():
                speak('Stopping text message mode.')
                try:
                    messaging_thread.join(timeout=1.0)
                except Exception:
                    pass
                messaging_thread = None
            else:
                speak('Text message mode stopped.')
        else:
            speak('Text message mode is not active.')
        return
    if any(x in command for x in ["enable text message", "enable text", "text message"]):
        # start text mode in a background thread
        if not _focus_whatsapp_window():
            speak('WhatsApp is not open. Please open it first.')
        else:
            # prevent multiple workers from being started
            if messaging_mode == 'text' and messaging_thread and messaging_thread.is_alive():
                speak('Text message mode is already active.')
            else:
                messaging_mode = 'text'
                messaging_thread = threading.Thread(target=_text_message_worker, daemon=True)
                messaging_thread.start()
        return
    # (voice message feature removed) - use text mode or request voice-message upload if needed
    elif "forward message" in command:
        # forward last assistant-sent text message to another contact
        if not last_message_text:
            speak('There is no message to forward.')
            return
        speak('Who should I forward the message to?')
        target = get_voice_command()
        if not target:
            speak('No target provided.')
            return
        # open target chat and send last_message_text
        if _search_contact_and_prepare(target):
            if _open_top_search_result():
                time.sleep(0.6)
                _type_and_send_text_app(last_message_text)
                speak('Message forwarded.')
                return
        speak('Could not forward message to the requested contact.')
    elif any(x in command for x in ["enable dark mode", "enable dark", "dark mode", "turn on dark mode", "turn on dark"]):
        if set_windows_theme(dark=True):
            speak("Dark mode enabled.")
        else:
            speak("Could not enable dark mode.")
    elif any(x in command for x in ["enable light mode", "enable light", "light mode", "turn on light mode", "turn on light"]):
        if set_windows_theme(dark=False):
            speak("Light mode enabled.")
        else:
            speak("Could not enable light mode.")
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
