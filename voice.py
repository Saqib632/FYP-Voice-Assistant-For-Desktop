import os
import difflib
import time
import socket
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
import json
from datetime import datetime
import winreg
import ctypes
from PIL import ImageGrab
import pygetwindow as gw

try:
    import wmi
    HAS_WMI = True
except Exception:
    wmi = None
    HAS_WMI = False

try:
    import pytesseract
    HAS_PYTESSERACT = True
except Exception:
    HAS_PYTESSERACT = False

# ========== GLOBAL VARIABLES FOR OFFLINE-FIRST MODE ==========
OFFLINE_MODE = False  # Will be auto-set at startup
search_mode_enabled = False
recording_thread = None

# ========== INTERNET CONNECTIVITY CHECKER ==========

def check_internet_connection():
    """
    Check if internet is available by attempting to ping Google DNS (8.8.8.8).
    Returns True if internet is available, False otherwise.
    """
    try:
        # Try to resolve a DNS name and connect to port 53 (DNS) or port 80 (HTTP)
        socket.create_connection(("8.8.8.8", 53), timeout=2)
        return True
    except (socket.timeout, socket.error, OSError):
        pass
    
    try:
        # Fallback: try to connect to Google DNS on HTTP port
        socket.create_connection(("8.8.8.8", 80), timeout=2)
        return True
    except (socket.timeout, socket.error, OSError):
        pass
    
    return False

# ========== VOSK OFFLINE SPEECH RECOGNITION SETUP ==========

_vosk_model = None
_VOSK_MODEL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "vosk-model-small-en-us-0.15")

# ========== VOSK GRAMMAR (constrains recognition to known commands) ==========
# This is the BIGGEST accuracy improvement: instead of matching against the
# entire English language, Vosk only considers these phrases.
_VOSK_GRAMMAR = json.dumps([
    # Wake word
    "nova", "hey nova",
    # Browser & Web (split brand names into words the model knows)
    "open you tube", "open google", "open face book", "open get hub",
    "open chrome", "close chrome", "close google chrome", "close browser",
    "close the chrome", "close the browser", "exit chrome", "quit chrome",
    # Tabs
    "new tab", "open new tab", "open a new tab", "open tab", "create tab",
    "next tab", "switch tab", "go to next tab",
    "previous tab", "go to previous tab", "back tab", "last tab",
    "close tab", "exit tab",
    # WhatsApp (split into words the model knows)
    "open whats app", "whats app", "close whats app", "exit whats app",
    "enable text message", "enable text", "text message",
    "disable text message", "disable text", "disable text mode",
    "stop text mode", "exit text mode", "stop text message",
    "forward message",
    # Volume
    "increase volume", "decrease volume", "volume up", "volume down",
    "mute", "un mute", "mute volume",
    # Brightness
    "increase brightness", "decrease brightness",
    "brightness up", "brightness down",
    # Theme
    "enable dark mode", "enable light mode", "dark mode", "light mode",
    "turn on dark mode", "turn on light mode",
    "disable dark mode", "disable light mode",
    "turn off dark mode", "turn off light mode",
    "enable dark", "enable light", "disable dark", "disable light",
    # WiFi (split into words the model knows)
    "turn on wi fi", "turn off wi fi", "enable wi fi", "disable wi fi",
    "wi fi on", "wi fi off",
    # Search
    "search", "search for", "enable search mode", "disable search mode",
    "google search",
    # Apps (split compound words)
    "open calculator", "close calculator",
    "open note pad", "close note pad",
    # Folders
    "open downloads", "open music", "open videos", "open desktop", "open documents",
    "close downloads", "close music", "close videos", "close desktop", "close documents",
    # Screen (split compound words)
    "start screen recording", "stop screen recording",
    "take a screen shot", "take screen shot", "capture screen",
    # System
    "exit", "stop", "cancel", "quit",
    # Weather & AI queries
    "what is", "who is", "weather", "weather today",
    # Catch-all for unrecognized words
    "[unk]"
])

# ========== KNOWN COMMANDS LIST (for fuzzy matching) ==========
_KNOWN_COMMANDS = [
    "nova",
    "open youtube", "open google", "open facebook", "open github",
    "open chrome", "close chrome", "close google chrome", "close browser",
    "new tab", "open new tab", "next tab", "previous tab", "close tab",
    "open whatsapp", "close whatsapp",
    "enable text message", "disable text message", "stop text mode",
    "forward message",
    "increase volume", "decrease volume", "mute", "unmute",
    "increase brightness", "decrease brightness",
    "enable dark mode", "enable light mode", "dark mode", "light mode",
    "disable dark mode", "disable light mode",
    "turn on wifi", "turn off wifi", "enable wifi", "disable wifi",
    "search", "enable search mode", "disable search mode",
    "open calculator", "close calculator",
    "open notepad", "close notepad",
    "open downloads", "open music", "open videos", "open desktop", "open documents",
    "close downloads", "close music", "close videos", "close desktop", "close documents",
    "start screen recording", "stop screen recording",
    "take a screenshot", "capture screen",
    "exit", "stop", "cancel",
]

def _fuzzy_match_command(text, threshold=0.55):
    """Find the closest known command using fuzzy matching.
    Returns the best match if similarity >= threshold, otherwise returns original text.
    This catches near-misses like 'close crome' -> 'close chrome'.
    """
    if not text:
        return text
    # First check for exact substring matches
    for cmd in _KNOWN_COMMANDS:
        if cmd in text or text in cmd:
            return text  # already a valid match, don't change
    # Try fuzzy matching
    matches = difflib.get_close_matches(text, _KNOWN_COMMANDS, n=1, cutoff=threshold)
    if matches:
        print(f"[Fuzzy] '{text}' -> '{matches[0]}'")
        return matches[0]
    return text

def _download_vosk_model():
    """
    Download the Vosk small English model (~40 MB) on first run.
    Only called when internet is available and model is missing.
    """
    import urllib.request
    import zipfile
    model_url = "https://alphacephei.com/vosk/models/vosk-model-small-en-us-0.15.zip"
    base_dir  = os.path.dirname(os.path.abspath(__file__))
    zip_path  = os.path.join(base_dir, "vosk-model-small-en-us-0.15.zip")
    try:
        print("Downloading Vosk speech model (~40 MB) - please wait...")
        urllib.request.urlretrieve(model_url, zip_path)
        print("Extracting model...")
        with zipfile.ZipFile(zip_path, 'r') as zf:
            zf.extractall(base_dir)
        os.remove(zip_path)
        print("[OK] Vosk model downloaded and ready.")
        return True
    except Exception as e:
        print(f"[!] Vosk model download failed: {e}")
        if os.path.exists(zip_path):
            try:
                os.remove(zip_path)
            except Exception:
                pass
        return False

def _init_vosk_model():
    """
    Load the Vosk offline model.  Auto-download if missing and internet available.
    Returns True if the model is ready, False otherwise.
    """
    global _vosk_model
    if _vosk_model is not None:
        return True
    try:
        import vosk
        vosk.SetLogLevel(-1)  # suppress noisy C++ logs
        if not os.path.exists(_VOSK_MODEL_PATH):
            if check_internet_connection():
                print("[!] Vosk model not found. Downloading automatically...")
                if not _download_vosk_model():
                    print("[!] Could not download Vosk model.")
                    return False
            else:
                print("[!] Vosk model not found and no internet available.")
                print("    To enable voice commands: connect to internet once to auto-download the model (40 MB).")
                print("    Until then, all commands can be typed at the prompt.")
                return False
        _vosk_model = vosk.Model(_VOSK_MODEL_PATH)
        print("[OK] Vosk offline speech model loaded.")
        return True
    except ImportError:
        print("[!] vosk package not installed. Run: pip install vosk")
        return False
    except Exception as e:
        print(f"[!] Vosk model init failed: {e}")
        return False

def _recognize_with_vosk(audio, use_grammar=True):
    """
    Transcribe an sr.AudioData object using the loaded Vosk model.
    Uses grammar-constrained recognition by default for command accuracy.
    Falls back to unconstrained recognition when use_grammar=False
    (for free-form input like contact names or search queries).
    Raises sr.UnknownValueError when nothing was understood.
    """
    import vosk
    # Convert to 16-kHz 16-bit PCM that Vosk expects
    wav_bytes = audio.get_wav_data(convert_rate=16000, convert_width=2)

    # Create recognizer: grammar-constrained or free-form
    if use_grammar:
        rec = vosk.KaldiRecognizer(_vosk_model, 16000, _VOSK_GRAMMAR)
    else:
        rec = vosk.KaldiRecognizer(_vosk_model, 16000)

    # Stream audio in chunks for better accuracy
    # (Vosk's internal language model tracks context better with streaming)
    chunk_size = 4000
    for i in range(0, len(wav_bytes), chunk_size):
        rec.AcceptWaveform(wav_bytes[i:i + chunk_size])

    result = json.loads(rec.FinalResult())
    text = result.get("text", "").strip()

    # Filter out Vosk's [unk] placeholder
    text = text.replace("[unk]", "").strip()

    if not text:
        raise sr.UnknownValueError()
    return text

def is_chrome_running():
    """Checks if Chrome is running"""
    for proc in os.popen('tasklist').readlines():
        if 'chrome.exe' in proc:
            return True
    return False

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
                cmd = get_voice_command(free_form=True)
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

    # ---- Rejoin split words from Vosk grammar ----
    # The Vosk small model doesn't know compound/brand words, so the grammar
    # uses split forms. We rejoin them here to match the command handler.
    normalized = normalized.replace("you tube", "youtube")
    normalized = normalized.replace("face book", "facebook")
    normalized = normalized.replace("get hub", "github")
    normalized = normalized.replace("whats app", "whatsapp")
    normalized = normalized.replace("note pad", "notepad")
    normalized = normalized.replace("screen shot", "screenshot")
    normalized = normalized.replace("un mute", "unmute")
    normalized = normalized.replace("wi-fi", "wifi")
    normalized = normalized.replace("wi fi", "wifi")

    # ---- Phonetic corrections for "nova" wake-word mishearings ----
    _nova_phonetics = [
        "no one", "no wan", "no won", "no want", "know one",
        "wrong", "ron", "rong", "grown",
        "no", "know", "known", "now",
        "noba", "novia", "nava", "novah", "no va", "nova",
        "mova", "moved", "move", "lover", "over",
        "noaa", "nope", "nota", "norma", "noah"
    ]
    if normalized in _nova_phonetics:
        normalized = "nova"

    # ---- Phonetic corrections for "open" mishearings ----
    _open_phonetics = [
        "upon", "often", "hoping", "oven", "oh pen",
        "oh been", "oben", "opan", "opin"
    ]
    _open_targets = [
        "youtube", "google", "facebook", "github", "chrome",
        "whatsapp", "calculator", "notepad",
        "downloads", "music", "videos", "desktop", "documents"
    ]
    for misspell in _open_phonetics:
        for target in _open_targets:
            normalized = normalized.replace(misspell + " " + target, "open " + target)

    # ---- Phonetic corrections for "close" mishearings ----
    _close_phonetics = [
        "laws", "clause", "claws", "clothes", "clues", "glows", "flows",
        "clos", "cloz", "clothe", "cloe", "lows", "lose", "closs", "cloes",
        "clouse", "kloz", "klose", "klos", "cloth", "glows", "blows",
        "plus", "class", "glass", "gross"
    ]
    _close_targets = [
        "chrome", "google chrome", "notepad", "whatsapp", "tab",
        "the app", "window", "app", "browser", "calculator"
    ]
    for misspell in _close_phonetics:
        for target in _close_targets:
            normalized = normalized.replace(misspell + " " + target, "close " + target)

    # ---- Phonetic corrections for "tab" mishearings ----
    _tab_phonetics = [
        "bab", "bad", "dab", "dad", "nab", "jab", "lab", "cab",
        "tap", "tag", "tat", "bat", "mat", "nat", "pat", "rat",
        "sad", "lad", "had", "tab"
    ]
    _tab_prefixes = ["new ", "next ", "previous ", "close ", "open new "]
    for prefix in _tab_prefixes:
        for misspell in _tab_phonetics:
            normalized = normalized.replace(prefix + misspell, prefix + "tab")
    _tab_open_variants = ["open a new " + m for m in _tab_phonetics] + \
                         ["open new " + m for m in _tab_phonetics]
    for variant in _tab_open_variants:
        normalized = normalized.replace(variant, "new tab")

    # ---- Phonetic corrections for "volume" mishearings ----
    _volume_phonetics = ["follow me", "fallen", "volume", "vellum", "villain", "volley"]
    for misspell in _volume_phonetics:
        if misspell == "volume":
            continue
        normalized = normalized.replace("increase " + misspell, "increase volume")
        normalized = normalized.replace("decrease " + misspell, "decrease volume")

    # ---- Phonetic corrections for "brightness" mishearings ----
    _brightness_phonetics = ["rightness", "bright mess", "bright ness", "brightest", "writeness"]
    for misspell in _brightness_phonetics:
        normalized = normalized.replace("increase " + misspell, "increase brightness")
        normalized = normalized.replace("decrease " + misspell, "decrease brightness")

    # ---- Phonetic corrections for "screenshot" mishearings ----
    _screenshot_phonetics = [
        "screen shot", "screens hot", "screen shut",
        "screen shirt", "green shot", "screen chart"
    ]
    for misspell in _screenshot_phonetics:
        normalized = normalized.replace("take a " + misspell, "take a screenshot")
        normalized = normalized.replace(misspell, "screenshot")

    # ---- Phonetic corrections for "recording" mishearings ----
    _recording_phonetics = ["record in", "report in", "recording", "regarding"]
    for misspell in _recording_phonetics:
        if misspell == "recording":
            continue
        normalized = normalized.replace("start screen " + misspell, "start screen recording")
        normalized = normalized.replace("stop screen " + misspell, "stop screen recording")

    # ---- Phonetic corrections for "mute/unmute" mishearings ----
    _mute_phonetics = ["meet", "mood", "moot", "moot", "muted"]
    if normalized in _mute_phonetics:
        normalized = "mute"
    _unmute_phonetics = ["on mute", "on meet", "unmuted", "and mute", "un mute"]
    if normalized in _unmute_phonetics:
        normalized = "unmute"

    # ---- Final: apply fuzzy matching to catch remaining near-misses ----
    normalized = _fuzzy_match_command(normalized)

    return normalized

def get_voice_command(prompt=None, free_form=False):
    """
    Offline-first speech recognition pipeline:
      Tier 1  - Vosk          (offline, grammar-constrained for commands)
      Tier 2  - PocketSphinx  (offline fallback when Vosk model not yet downloaded)
      Tier 3  - Google SR     (online fallback, only when internet available)

    When free_form=True, skips grammar constraints (for contact names,
    search queries, etc.)

    Recognition failures (could not understand) -> return None -> main loop retries.
    Text input is ONLY used when the microphone itself is unavailable.
    """
    global OFFLINE_MODE

    recognizer = sr.Recognizer()
    # Optimized audio capture settings for better command recognition
    recognizer.energy_threshold         = 1500   # lower = picks up quieter speech
    recognizer.dynamic_energy_threshold = True
    recognizer.pause_threshold          = 0.7    # faster response for short commands
    recognizer.phrase_threshold         = 0.2    # catch shorter utterances
    recognizer.non_speaking_duration    = 0.5    # faster end-of-phrase detection

    # ---------- capture audio ----------
    try:
        with sr.Microphone() as source:
            # Longer ambient noise calibration for a better noise floor
            recognizer.adjust_for_ambient_noise(source, duration=1.5)
            if prompt:
                speak(prompt)
            print("Listening...")
            # timeout=15 waits up to 15s for speech to start
            # phrase_time_limit=10 captures up to 10s of actual speech
            audio = recognizer.listen(source, timeout=15, phrase_time_limit=10)
    except sr.WaitTimeoutError:
        return None          # silence - loop back quietly
    except OSError:
        print("Microphone not found. Please type your command.")
        return get_text_input()
    except Exception as e:
        print(f"Audio capture error: {e}")
        return get_text_input()

    print("Processing your command...")

    # ---------- TIER 1: Vosk (offline, grammar-constrained) ----------
    if _vosk_model is not None:
        try:
            # First pass: grammar-constrained (unless free_form requested)
            text = _recognize_with_vosk(audio, use_grammar=not free_form)
            command = normalize_command(text)
            if command:
                print(f"Command: {command}")
                return command
        except sr.UnknownValueError:
            # Grammar-constrained pass failed; try unconstrained as fallback
            if not free_form:
                try:
                    text = _recognize_with_vosk(audio, use_grammar=False)
                    command = normalize_command(text)
                    if command:
                        print(f"Command (fallback): {command}")
                        return command
                except sr.UnknownValueError:
                    pass  # truly could not understand
                except Exception as e:
                    print(f"[Vosk fallback] Error: {e}")
        except Exception as e:
            print(f"[Vosk] Error: {e}")

    # ---------- TIER 2: PocketSphinx (offline fallback) ----------
    elif OFFLINE_MODE:
        try:
            text = recognizer.recognize_sphinx(audio)
            command = normalize_command(text)
            if command:
                print(f"Command: {command}")
                return command
        except sr.UnknownValueError:
            pass   # could not understand - loop back
        except Exception as e:
            print(f"[Offline] Recognition error: {e}")

    # ---------- TIER 3: Google SR (online only) ----------
    if not OFFLINE_MODE:
        try:
            text = recognizer.recognize_google(audio)
            command = normalize_command(text)
            if command:
                print(f"Command: {command}")
                return command
        except sr.RequestError:
            print("[Google SR] Unavailable - switching to offline mode.")
            OFFLINE_MODE = True
            speak("Internet lost. Switching to offline mode.")
        except sr.UnknownValueError:
            pass   # could not understand - loop back

    # Could not understand in any engine - return None so main loop retries
    return None

def get_text_input():
    """Text input fallback when speech recognition is unavailable or fails"""
    try:
        command = input("Type command: ").strip()
        if command:
            return normalize_command(command)
        return None
    except KeyboardInterrupt:
        return "exit"
    except EOFError:
        return "exit"

def ask_openrouter(question):
    """
    Query OpenRouter API for LLM-based responses.
    Only works when internet is available.
    Returns a response string or offline message.
    """
    global OFFLINE_MODE
    
    if OFFLINE_MODE:
        offline_msg = "This feature requires internet connection. Please enable internet to use this feature."
        speak(offline_msg)
        return offline_msg
    
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json"
    }
    data = {
        "model": "mistralai/mistral-7b-instruct",
        "messages": [{"role": "user", "content": question}]
    }
    try:
        response = requests.post(API_URL, headers=headers, json=data, timeout=10)
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"]
    except requests.exceptions.RequestException as e:
        print(f"API request failed: {e}")
        # Mark offline mode if internet failed
        OFFLINE_MODE = True
        offline_msg = "Internet connection lost. This feature is unavailable in offline mode."
        speak(offline_msg)
        return offline_msg
    except Exception as e:
        return f"Failed to get a response: {e}"

def open_item(path):
    try:
        os.startfile(path)
        speak(f"Opening {path}")
    except Exception as e:
        speak(f"Unable to open: {e}")

def open_website(url):
    """
    Open a website in Chrome.
    Only works when internet is available.
    """
    global OFFLINE_MODE
    
    if OFFLINE_MODE:
        offline_msg = f"Internet not available. Cannot open {url.replace('.com', '')}."
        speak(offline_msg)
        return False
    
    try:
        chrome_path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
        webbrowser.register('chrome', None, webbrowser.BackgroundBrowser(chrome_path))
        webbrowser.get('chrome').open(f"https://{url}")
        speak(f"Opening {url.replace('.com','')}")
        return True
    except Exception as e:
        speak(f"Failed to open {url}: {e}")
        return False

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
        from comtypes import CLSCTX_ALL
        from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
        from ctypes import cast, POINTER
        
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(
            IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        
        if action == "increase":
            current = volume.GetMasterVolumeLevelScalar()
            volume.SetMasterVolumeLevelScalar(min(current + 0.1, 1.0), None)
            speak("Volume increased.")
        elif action == "decrease":
            current = volume.GetMasterVolumeLevelScalar()
            volume.SetMasterVolumeLevelScalar(max(current - 0.1, 0.0), None)
            speak("Volume decreased.")
        elif action == "mute":
            volume.SetMute(1, None)
            speak("Volume muted.")
        elif action == "unmute":
            is_muted = volume.GetMute()
            if is_muted:
                volume.SetMute(0, None)
                speak("Volume unmuted.")
            else:
                speak("Already unmuted.")
    except AttributeError as e:
        # Fallback to keyboard shortcuts using pyautogui
        try:
            if action == "mute":
                pyautogui.press("volumemute")
                speak("Volume muted.")
            elif action == "unmute":
                pyautogui.press("volumemute")
                speak("Volume unmuted.")
            elif action == "increase":
                for _ in range(2):
                    pyautogui.press("volumeup")
                speak("Volume increased.")
            elif action == "decrease":
                for _ in range(2):
                    pyautogui.press("volumedown")
                speak("Volume decreased.")
        except Exception as fallback_error:
            speak("Volume control not available.")
    except Exception as e:
        speak(f"Volume error: {e}")

def change_brightness(action):
    try:
        if not HAS_WMI:
            speak("Brightness control requires the optional 'wmi' package.")
            return
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
        
        # Use PowerShell with elevated privileges to run netsh command
        ps_command = f"Start-Process -Verb RunAs -FilePath 'netsh' -ArgumentList 'interface set interface name=\"{interface_name}\" admin={state}' -Wait -WindowStyle Hidden"
        
        # Try PowerShell approach first (requires elevation at Electron level)
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_command],
            capture_output=True,
            text=True,
            shell=True
        )

        if result.returncode == 0:
            status = 'turned on' if enabled else 'turned off'
            speak(f"Wi-Fi {status}.")
            print(f"[SUCCESS] Wi-Fi {status}")
        else:
            # Fallback: Try direct netsh command (may work if Electron is admin)
            command = f'netsh interface set interface name="{interface_name}" admin={state}'
            result = subprocess.run(command, shell=True, capture_output=True, text=True)
            
            if result.returncode == 0:
                status = 'turned on' if enabled else 'turned off'
                speak(f"Wi-Fi {status}.")
                print(f"[SUCCESS] Wi-Fi {status}")
            else:
                error_msg = result.stderr if result.stderr else "Access Denied - Admin Required"
                print(f"[ERROR] WiFi command failed: {error_msg}")
                speak(f"Failed to {'enable' if enabled else 'disable'} Wi-Fi. The application must be running as administrator.")
    except Exception as e:
        print(f"[ERROR] Wi-Fi control exception: {e}")
        speak(f"Wi-Fi control error: {e}")

# ========== GOOGLE SEARCH MODE ==========

def google_search_mode():
    """
    Enable Google search mode (online-only feature).
    Returns early if internet is not available.
    """
    global OFFLINE_MODE
    
    if OFFLINE_MODE:
        speak("Internet not available. Cannot access Google Search.")
        return
    
    speak("What would you like to search for?")
    while True:
        query = get_voice_command(free_form=True)
        if query and any(x in query for x in ["exit", "stop", "cancel"]):
            speak("Exiting search.")
            break
        elif query:
            try:
                chrome_path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
                webbrowser.register('chrome', None, webbrowser.BackgroundBrowser(chrome_path))
                webbrowser.get('chrome').open(f"https://www.google.com/search?q={query}")
                speak(f"Searching for {query}")
            except Exception as e:
                speak(f"Search failed: {e}")

# ========== CHROME TABS CONTROL ==========

def manage_chrome_tabs(action):
    try:
        if not is_chrome_running():
            speak("Chrome browser is not open. Please open Chrome first to manage tabs.")
            return
        # Bring Chrome to the foreground
        chrome_windows = [w for w in gw.getWindowsWithTitle('Google Chrome') if w.visible]
        if chrome_windows:
            chrome_win = chrome_windows[0]
            # Try pygetwindow activate first; fall back to ctypes on Windows pipe error
            try:
                chrome_win.activate()
            except Exception:
                try:
                    hwnd = chrome_win._hWnd
                    ctypes.windll.user32.ShowWindow(hwnd, 9)   # SW_RESTORE = 9
                    ctypes.windll.user32.SetForegroundWindow(hwnd)
                except Exception:
                    pass
            time.sleep(0.3)  # Give time for window to focus
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
    global search_mode_enabled, OFFLINE_MODE
    global awaiting_contact_search, awaiting_contact_selection, _whatsapp_search_query, whatsapp_current_chat, messaging_mode, messaging_thread

    # ---- Wake word: bring Nova window to foreground ----
    # When the user says "nova" (or close variants), signal the Electron
    # main process to restore and focus the window.
    _aivon_variants = [
        "nova", "no va", "nова", "nova assistant",
        "maven", "noba", "novah", "novia", "nava",
        # common mishearings observed in logs:
        "no one", "no wan", "no won", "know one", "no want",
        "wrong", "ron", "rong",
        "no", "know", "known", "now",
        "mova", "move", "moved"
    ]
    if any(command.strip() == v for v in _aivon_variants):
        print("[FOCUS_WINDOW]", flush=True)
        speak("Yes, I am here.")
        return

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
        if OFFLINE_MODE:
            speak("Internet not available. Search mode requires internet connection.")
            return
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
        elif command and not OFFLINE_MODE:
            try:
                chrome_path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
                webbrowser.register('chrome', None, webbrowser.BackgroundBrowser(chrome_path))
                webbrowser.get('chrome').open(f"https://www.google.com/search?q={command}")
                speak(f"Searching for {command}")
            except Exception as e:
                speak(f"Search failed: {e}")
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
    elif any(x in command for x in [
        "close chrome", "close google chrome", "close the chrome",
        "shut chrome", "shut down chrome", "exit chrome", "quit chrome",
        "kill chrome", "stop chrome", "close browser", "close the browser",
        "exit browser", "quit browser"
    ]):
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
        target = get_voice_command(free_form=True)
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
    # Disable/Turn-off handlers for dark/light mode (more specific patterns checked before enable)
    elif any(x in command for x in ["disable dark mode", "disable dark", "turn off dark mode", "turn off dark"]):
        if set_windows_theme(dark=False):
            speak("Dark mode disabled.")
        else:
            speak("Could not disable dark mode.")
    elif any(x in command for x in ["disable light mode", "disable light", "turn off light mode", "turn off light"]):
        if set_windows_theme(dark=True):
            speak("Light mode disabled.")
        else:
            speak("Could not disable light mode.")
    # Enable/Turn-on handlers for dark/light mode (checked after disable patterns)
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
    elif any(x in command for x in [
        "new tab", "open new tab", "open a new tab", "a new tab",
        "open tab", "create tab", "create new tab"
    ]):
        manage_chrome_tabs("new")
    elif any(x in command for x in [
        "next tab", "switch tab", "tab next", "go to next tab",
        "move to next tab", "forward tab"
    ]):
        manage_chrome_tabs("next")
    elif any(x in command for x in [
        "previous tab", "prev tab", "tab previous", "go to previous tab",
        "move to previous tab", "back tab", "go back tab", "last tab"
    ]):
        manage_chrome_tabs("previous")
    elif any(x in command for x in [
        "close tab", "shut tab", "exit tab", "remove tab", "delete tab"
    ]):
        manage_chrome_tabs("close")
    elif "search for" in command or "what is the weather" in command or "weather today" in command or command.strip() == "weather":
        # Weather and search are online-only features
        if OFFLINE_MODE:
            speak("Internet not available. Cannot search or get weather information.")
            return
        
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
        
        try:
            speak(f"Searching for {query}")
            chrome_path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
            webbrowser.register('chrome', None, webbrowser.BackgroundBrowser(chrome_path))
            webbrowser.get('chrome').open(f"https://www.google.com/search?q={query}")
        except Exception as e:
            speak(f"Search failed: {e}")
    elif "what is" in command or "who is" in command or "define" in command:
        # LLM queries are online-only
        if OFFLINE_MODE:
            speak("Internet not available. This feature requires internet connection.")
            return
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
    global OFFLINE_MODE

    print("="*60)
    print("       VOICE ASSISTANT  -  OFFLINE-FIRST")
    print("="*60)

    # 1. Internet check
    print("[1/2] Checking internet connectivity...")
    OFFLINE_MODE = not check_internet_connection()

    # 2. Vosk model init  (auto-download if internet available)
    print("[2/2] Loading offline speech model...")
    vosk_ready = _init_vosk_model()

    # ---- status report ----
    print("="*60)
    if OFFLINE_MODE:
        print("  MODE : OFFLINE")
        print("  VOICE: " + ("Vosk (offline)" if vosk_ready else "Text input only"))
        print("  APPS : volume / brightness / wifi / theme / apps / tabs")
        print("  OFF  : web search / weather / LLM (need internet)")
        speak("Offline mode active. System commands are ready.")
    else:
        print("  MODE : ONLINE")
        print("  VOICE: " + ("Vosk (offline) + Google fallback" if vosk_ready else "Google Speech Recognition"))
        print("  ALL  : system automation + search + weather + LLM")
        speak("Online mode. All features are available.")
    print("="*60)
    print("Speak a command clearly, or type it when prompted.")
    print("Press Ctrl+C to exit.")
    print()

    while True:
        try:
            command = get_voice_command()
            if command:
                handle_command(command)
            # None = silence/timeout -> loop back silently
        except KeyboardInterrupt:
            print("\nExiting...")
            speak("Goodbye!")
            break
        except Exception as e:
            print(f"Error: {e}")
            time.sleep(0.5)

if __name__ == "__main__":
    main()
