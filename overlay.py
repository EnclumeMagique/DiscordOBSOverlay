import time
import ctypes
from pywinauto import Application
from obswebsocket import obsws, requests
from concurrent.futures import ThreadPoolExecutor
import tkinter as tk
import threading
from threading import Thread
import comtypes
import comtypes.client
import comtypes.stream
from flask import Flask, request
import os

flask_app = Flask(__name__)

@flask_app.route('/toggle_inverse', methods=['POST'])
def toggle_inverse():
    toggle_reverse_overlay()
    return 'Inverse Mode Toggled', 200

@flask_app.route('/toggle_overlay', methods=['POST'])
def toggle_overlay():
    toggle_overlay()
    return 'Overlay Toggled', 200

@flask_app.route('/force_reverse', methods=['POST'])
def force_reverse():
    global reverse_overlay_active
    if not reverse_overlay_active:
        toggle_reverse_overlay()
    return 'Reverse Mode Forced ON', 200

@flask_app.route('/force_reverse_off', methods=['POST'])
def force_reverse_off():
    global reverse_overlay_active
    if reverse_overlay_active:
        toggle_reverse_overlay()
    return 'Reverse Mode Forced OFF', 200

def start_flask():
    from waitress import serve
    serve(flask_app, host='0.0.0.0', port=5001)

# Set up the Win32 API call to bring the window to the foreground
SetForegroundWindow = ctypes.windll.user32.SetForegroundWindow
FindWindow = ctypes.windll.user32.FindWindowW

def find_discord_window():
    app = Application(backend="uia").connect(title_re="^(?!.*Microsoft Edge).*(Discord).*", found_index=0)
    windows = app.windows()
    for win in windows:
        if "Discord" in win.window_text() and "Microsoft Edge" not in win.window_text():
            if has_mute_button(win):
                return win
    return None

# Function to find the mute button in a window
def has_mute_button(window):
    if window is None:
        return False
    try:
        mute_button = find_button(window, "Mute")
        return mute_button is not None
    except Exception as e:
        print(f"Error checking for mute button: {e}")
        return False

# Function to find a button in a window
def find_button(window, title):
    if window is None:
        return None
    try:
        for child in window.descendants(control_type="Button"):
            if child.window_text() == title:
                return child
        return None
    except Exception as e:
        print(f"Error finding button '{title}': {e}")
        return None

# Retry mechanism to find the Discord window
discord_window = None
for _ in range(5):  # Retry up to 5 times
    discord_window = find_discord_window()
    if discord_window:
        break
    print("Retrying to find Discord window...")
    time.sleep(2)

if discord_window:
    hwnd = FindWindow(None, discord_window.window_text())
    SetForegroundWindow(hwnd)
    print("Discord window brought to the foreground.")
else:
    print("No matching Discord window with mute button found.")

# OBS WebSocket settings
OBS_HOST = "127.0.0.1"  # or "localhost"
OBS_PORT = 4455  # Ensure this matches the port in your OBS WebSocket settings
OBS_PASSWORD = ""  # Leave empty since authentication is disabled
ws = obsws(OBS_HOST, OBS_PORT, OBS_PASSWORD)

SCENE_NAME = "Stream effects"  # Update this with your actual scene name

def get_scene_item_id(source_name):
    try:
        response = ws.call(requests.GetSceneItemList(sceneName=SCENE_NAME))
        scene_items = response.getSceneItems()  # Correct method to retrieve scene items
        for item in scene_items:
            if item["sourceName"] == source_name:
                return item["sceneItemId"]  # Correctly retrieve scene item ID
        print(f"Source {source_name} not found in scene {SCENE_NAME}.")
        return None
    except Exception as e:
        print(f"Failed to retrieve scene items for {SCENE_NAME}: {e}")
        return None

def set_visibility(source_name, visible):
    try:
        item_id = get_scene_item_id(source_name)
        if item_id is not None:
            ws.call(requests.SetSceneItemEnabled(sceneName=SCENE_NAME, sceneItemId=item_id, sceneItemEnabled=visible))
            print(f"Set {source_name} visibility to {visible}")
        else:
            print(f"Failed to set visibility for {source_name}: item not found.")
    except Exception as e:
        print(f"Failed to update OBS source {source_name}: {e}")

def check_mute_status(window):
    try:
        mute_button = find_button(window, "Mute")
        if mute_button:
            mute_state = mute_button.get_toggle_state()
            return mute_state
        else:
            print("Mute button not found.")
            return None
    except Exception as e:
        print(f"Discord mute info error: {e}")
        return None

def check_deafen_status(window):
    try:
        deafen_button = find_button(window, "Deafen")
        if deafen_button:
            deafen_state = deafen_button.get_toggle_state()
            return deafen_state
        else:
            print("Deafen button not found.")
            return None
    except Exception as e:
        print(f"Discord deafen info error: {e}")
        return None

def check_status():
    previous_mute_state = None
    previous_deafen_state = None
    previous_audio_info_state = None

    set_visibility("Discord audio info", False)
    set_visibility("Discord audio info error", False)
    set_visibility("Discord inverse unmute", False)

    with ThreadPoolExecutor(max_workers=2) as executor:
        while not stop_event.is_set():
            future_mute = executor.submit(check_mute_status, discord_window)
            future_deafen = executor.submit(check_deafen_status, discord_window)
            mute_state = future_mute.result()
            deafen_state = future_deafen.result()

            if mute_state is not None and deafen_state is not None:
                if reverse_overlay_active:
                    if mute_state == 0 and deafen_state == 0:
                        set_visibility("Discord audio info", True)
                    else:
                        set_visibility("Discord audio info", False)
                    set_visibility("Discord Mute", False)
                    set_visibility("Discord Deafen", False)
                else:
                    if mute_state != previous_mute_state:
                        set_visibility("Discord Mute", mute_state == 1)
                        set_visibility("Discord inverse unmute", False)
                        previous_mute_state = mute_state

                    if deafen_state != previous_deafen_state:
                        set_visibility("Discord Deafen", deafen_state == 1)
                        previous_deafen_state = deafen_state

                    audio_info_state = mute_state == 1 or deafen_state == 1
                    if audio_info_state != previous_audio_info_state:
                        set_visibility("Discord audio info", audio_info_state)
                        previous_audio_info_state = audio_info_state

            time.sleep(0.1)  # Check every 0.1 second

def check_reverse_status():
    previous_mute_state = None
    previous_deafen_state = None

    set_visibility("Discord inverse unmute", False)

    with ThreadPoolExecutor(max_workers=2) as executor:
        while not reverse_stop_event.is_set():
            future_mute = executor.submit(check_mute_status, discord_window)
            future_deafen = executor.submit(check_deafen_status, discord_window)
            mute_state = future_mute.result()
            deafen_state = future_deafen.result()

            if mute_state is not None and deafen_state is not None:
                if mute_state != previous_mute_state or deafen_state != previous_deafen_state:
                    if mute_state == 0 and deafen_state == 0:
                        set_visibility("Discord audio info", True)
                        set_visibility("Discord inverse unmute", True)
                    else:
                        set_visibility("Discord audio info", False)
                        set_visibility("Discord inverse unmute", False)
                    previous_mute_state = mute_state
                    previous_deafen_state = deafen_state

            time.sleep(0.1)  # Check every 0.1 second

def toggle_overlay():
    global overlay_active, reverse_overlay_active
    overlay_active = not overlay_active
    status_label.config(text=f"Overlay is {'ON' if overlay_active else 'OFF'}")

    if overlay_active:
        stop_event.clear()
        status_thread = Thread(target=check_status)
        status_thread.start()
    else:
        stop_event.set()
        reverse_overlay_active = False  # Disable reverse overlay if main overlay is turned off
        reverse_status_label.config(text="Reverse Overlay is OFF")
        reverse_stop_event.set()

        # Hide all overlay elements
        set_visibility("Discord Mute", False)
        set_visibility("Discord Deafen", False)
        set_visibility("Discord audio info", False)
        set_visibility("Discord audio info error", False)
        set_visibility("Discord inverse unmute", False)  # Hide inverse unmute

def toggle_reverse_overlay():
    global reverse_overlay_active
    if overlay_active:  # Only allow toggling reverse overlay if main overlay is active
        reverse_overlay_active = not reverse_overlay_active
        reverse_status_label.config(text=f"Reverse Overlay is {'ON' if reverse_overlay_active else 'OFF'}")

        # Refresh the visibility states when toggling reverse mode
        if reverse_overlay_active:
            stop_event.set()
            reverse_stop_event.clear()
            reset_visibility_for_reverse_mode()
            reverse_status_thread = Thread(target=check_reverse_status)
            reverse_status_thread.start()
        else:
            reverse_stop_event.set()
            stop_event.clear()
            reset_visibility_for_regular_mode()
            status_thread = Thread(target=check_status)
            status_thread.start()
    else:
        print("Cannot enable reverse overlay when the main overlay is off.")

def reset_visibility_for_reverse_mode():
    """Reset visibility settings for reverse mode."""
    set_visibility("Discord Mute", False)
    set_visibility("Discord Deafen", False)
    set_visibility("Discord inverse unmute", False)
    set_visibility("Discord audio info", False)  # Start with all hidden

def reset_visibility_for_regular_mode():
    """Reset visibility settings for regular mode."""
    set_visibility("Discord Mute", False)
    set_visibility("Discord Deafen", False)
    set_visibility("Discord inverse unmute", False)
    set_visibility("Discord audio info", False)
    set_visibility("Discord audio info error", False)  # Ensure errors are hidden

def on_closing():
    stop_event.set()
    reverse_stop_event.set()
    set_visibility("Discord Mute", False)
    set_visibility("Discord Deafen", False)
    set_visibility("Discord audio info", False)
    set_visibility("Discord audio info error", True)
    set_visibility("Discord inverse unmute", False)  # Hide inverse unmute
    ws.disconnect()
    root.destroy()
    os._exit(0)  # Forcefully terminate the script

try:
    ws.connect()
    print("Connected to OBS WebSocket server.")

    stop_event = threading.Event()
    reverse_stop_event = threading.Event()

    overlay_active = True
    reverse_overlay_active = False
    status_thread = Thread(target=check_status)
    status_thread.start()

    threading.Thread(target=start_flask).start()

    root = tk.Tk()
    root.title("OBS Overlay Control")
    root.geometry("300x250")

    status_label = tk.Label(root, text="Overlay is ON")
    status_label.pack(pady=10)

    toggle_button = tk.Button(root, text="Toggle Overlay", command=toggle_overlay)
    toggle_button.pack(pady=10)

    reverse_status_label = tk.Label(root, text="Reverse Overlay is OFF")
    reverse_status_label.pack(pady=10)

    reverse_toggle_button = tk.Button(root, text="Toggle Reverse Overlay", command=toggle_reverse_overlay)
    reverse_toggle_button.pack(pady=10)

    exit_button = tk.Button(root, text="Exit", command=on_closing)
    exit_button.pack(pady=10)

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()

    status_thread.join()

except Exception as e:
    print(f"Failed to connect to OBS WebSocket server: {e}")
    ws.disconnect()
    os._exit(1)  # Forcefully terminate the script in case of an exception

# export command: python -m PyInstaller --onefile --noconsole overlay.py