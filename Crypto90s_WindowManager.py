import tkinter as tk
from tkinter import messagebox
import psutil
import pygetwindow as gw
import pickle
import subprocess
import time
import os
from screeninfo import get_monitors
import win32process
import win32gui
import win32con
import webbrowser
import pygetwindow as gw
import sys
import argparse
from pygetwindow import PyGetWindowException


import ctypes
from ctypes import wintypes

current_version = "v0.1.4"


def parse_args():
    parser = argparse.ArgumentParser(description="Crypto90's WindowManager Preset Selector")
    parser.add_argument("--preset", type=int, default=1, help="Preset number to load (default: 1)")
    return parser.parse_args()

# Class for window state
class WindowState:
    def __init__(self, process_name, position, size, process_path=None, url=None, path=None, minimized=False, launcher_override=None):
        self.process_name = process_name
        self.position = position
        self.size = size
        self.process_path = process_path
        self.url = url
        self.path = path
        self.minimized = minimized
        self.launcher_override = launcher_override

# Load/save state
def load_window_states(filename="window_states_1.pkl"):
    # Load the window states from the preset-specific file
    try:
        with open(filename, "rb") as f:
            data = pickle.load(f)
            # Assuming the new format is a dictionary with 'window_states' and 'config'
            if isinstance(data, dict):
                window_states = data.get("window_states", {})
                config = data.get("config", {})
                return window_states, config
    except (FileNotFoundError, EOFError):
        return {}, {}  # Return empty states and config on failure


def save_window_states(window_states, config=None, filename="window_states_1.pkl"):
    data = {
        "window_states": window_states,
        "config": config or {}
    }
    with open(filename, "wb") as f:
        pickle.dump(data, f)


def get_monitor_for_window(win, monitors):
    hwnd = win._hWnd
    try:
        # Use normal (restored) position if minimized
        if win32gui.IsIconic(hwnd):
            placement = win32gui.GetWindowPlacement(hwnd)
            # placement[4] is the normal (restored) position: (left, top, right, bottom)
            normal_pos = placement[4]
            wx = (normal_pos[0] + normal_pos[2]) // 2
            wy = (normal_pos[1] + normal_pos[3]) // 2
        else:
            wx = win.left + win.width // 2
            wy = win.top + win.height // 2

        for monitor in monitors:
            if monitor.x <= wx <= monitor.x + monitor.width and monitor.y <= wy <= monitor.y + monitor.height:
                return monitor

    except Exception as e:
        print(f"Error determining monitor for window '{win.title}': {e}")

    return None


def get_process_name_for_window(win):
    try:
        _, pid = win32process.GetWindowThreadProcessId(win._hWnd)
        return psutil.Process(pid).name()
    except Exception:
        return None

def get_process_path_for_window(win):
    try:
        _, pid = win32process.GetWindowThreadProcessId(win._hWnd)
        return psutil.Process(pid).exe()
    except Exception:
        return None

def is_process_running(process_path):
    for proc in psutil.process_iter(['exe']):
        try:
            if proc.info['exe'] and os.path.normcase(proc.info['exe']) == os.path.normcase(process_path):
                return True
        except psutil.AccessDenied:
            continue
    return False

def get_visible_windows_grouped_by_monitor():
    grouped = {}
    monitors = get_monitors()
    for window in gw.getAllWindows():
        hWnd = window._hWnd
        
        process_name = get_process_name_for_window(window)
        if not process_name or process_name == "explorer.exe":
            continue
        if not win32gui.IsWindowEnabled(hWnd):
            continue

        # Include windows that are either visible or minimized
        if not (win32gui.IsWindowVisible(hWnd) or win32gui.IsIconic(hWnd)):
            continue

        if win32gui.GetWindow(hWnd, win32con.GW_OWNER):
            continue

        if win32gui.GetWindowTextLength(hWnd) == 0:
            continue

        

        monitor = get_monitor_for_window(window, monitors)
        if monitor:
            mon_id = f"{monitor.x}x{monitor.y} {monitor.width}x{monitor.height}"
            if mon_id not in grouped:
                grouped[mon_id] = []
            grouped[mon_id].append((window, process_name))

    return grouped


def get_process_path_from_hwnd(hwnd):
    """Get process executable path from window handle."""
    pid = wintypes.DWORD()
    ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    try:
        proc = psutil.Process(pid.value)
        return proc.exe()
    except Exception:
        return None

def is_uwp_window(window):
    hwnd = window._hWnd  # Should work with pygetwindow or similar wrappers
    path = get_process_path_from_hwnd(hwnd)
    if path and "WindowsApps" in path:
        return True
    return False



#SpotifyAB.SpotifyMusic_1.262.580.0_x64__zpdnekdrzrea0
#SpotifyAB.SpotifyMusic_zpdnekdrzrea0!Spotify
def get_uwp_app_name(package_family_name_substring):
    try:
        # Extract the identifier part from package_family_name_substring (after the last '__')
        identifier = package_family_name_substring.split('__')[-1]
        
        # Run PowerShell command to get all start apps with full AppIDs using Format-List
        process = subprocess.Popen(
            ["powershell", "-Command", "Get-StartApps | Format-List Name, AppID"],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=False
        )

        stdout, stderr = process.communicate()

        # Try decoding with utf-8, fallback to cp1252 if it fails
        try:
            lines = stdout.decode('utf-8').splitlines()
        except UnicodeDecodeError:
            lines = stdout.decode('cp1252', errors='ignore').splitlines()

        # Find the header line and parse the output
        name = None
        appid = None
        for line in lines:
            if line.startswith("Name"):
                name = line.split(":", 1)[1].strip()
            elif line.startswith("AppID"):
                appid = line.split(":", 1)[1].strip()
                if identifier.lower() in appid.lower():
                    return name  # or return (name, appid) if both are useful

        return None

    except subprocess.CalledProcessError as e:
        print("Error running PowerShell:", e)
        return None
    except Exception as e:
        print("An unexpected error occurred:", e)
        return None











        
class WindowManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Crypto90's WindowManager")
        
        args = parse_args()
        preset_number = args.preset
        filename = f"window_states_{preset_number}.pkl"
        
        self.window_states, self.config = load_window_states(filename)
        self.window_mapping = []
        
        
        
        
        if "auto_close" in self.config:
            self.auto_close_var = tk.BooleanVar(value=self.config.get("auto_close"))
        else:
            self.auto_close_var = tk.BooleanVar(value=False)
        # Load auto-close state if it exists
        
        # Set minimum width and height for the main window
        self.root.minsize(width=400, height=400)
        
        
        # Set dark background for the main window
        self.root.configure(bg="#1e1e1e")

        # Process listbox with scrollbar
        self.process_listbox_frame = tk.Frame(self.root, bg="#1e1e1e")
        self.process_listbox_frame.pack(pady=10, fill=tk.BOTH, expand=True)

        self.process_listbox = tk.Listbox(
            self.process_listbox_frame,
            selectmode=tk.MULTIPLE,
            bg="#2e2e2e",
            fg="white",
            selectbackground="darkgreen",
            highlightbackground="#444",
            relief=tk.FLAT
        )
        self.process_listbox_scrollbar = tk.Scrollbar(
            self.process_listbox_frame,
            orient="vertical",
            command=self.process_listbox.yview,
            troughcolor="#2e2e2e",
            bg="#555"
        )
        self.process_listbox.config(yscrollcommand=self.process_listbox_scrollbar.set)

        self.process_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.process_listbox_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Preset Management Frame
        preset_frame = tk.Frame(self.root, bg="#1e1e1e")
        preset_frame.pack(pady=5)

        self.current_preset = tk.StringVar(value="Preset 1")

        tk.Label(preset_frame, text="Presets:", fg="white", bg="#1e1e1e").pack(side="left")
        for i in range(1, 11):
            name = f"Preset {i}"
            cb = tk.Radiobutton(
                preset_frame, text=str(i), variable=self.current_preset, value=name,
                command=self.switch_preset, bg="#1e1e1e", fg="white", selectcolor="#2e2e2e"
            )
            cb.pack(side="left", padx=2)
        
        # Auto-close checkbox
        self.auto_close_var = tk.BooleanVar()
        self.auto_close_checkbox = tk.Checkbutton(
            self.root, text="Auto-close after order", variable=self.auto_close_var,
            bg="#1e1e1e", fg="white", selectcolor="#2e2e2e"
        )
        self.auto_close_checkbox.pack()
        
        # Button frame
        button_frame = tk.Frame(self.root, bg="#1e1e1e")
        button_frame.pack(pady=10)

        tk.Button(button_frame, text="Refresh Process List", bg="#2980b9", fg="white", command=self.refresh_window_list).pack(side="left", padx=5)
        tk.Button(button_frame, text="Save Window Positions", bg="#7f8c8d", fg="white", command=self.save_window_positions).pack(side="left", padx=5)
        tk.Button(button_frame, text="Start, Resize and Order", bg="darkgreen", fg="white", command=self.stream_order).pack(side="left", padx=5)
        tk.Button(button_frame, text="Buy me a Coffee ☕", bg="#f39c12", fg="black", command=lambda: webbrowser.open("https://ko-fi.com/crypto90")).pack(side="left", padx=5)

        # Log box
        self.log_text_frame = tk.Frame(self.root, bg="#1e1e1e")
        self.log_text_frame.pack(pady=10, fill=tk.X, expand=False, side=tk.BOTTOM)

        self.log_text = tk.Text(
            self.log_text_frame,
            height=10,
            state=tk.DISABLED,
            bg="#1e1e1e",
            fg="#dcdcdc",
            insertbackground="white",
            highlightbackground="#444"
        )

        self.log_text_scrollbar = tk.Scrollbar(
            self.log_text_frame,
            orient="vertical",
            command=self.log_text.yview,
            troughcolor="#2e2e2e",
            bg="#555"
        )
        self.log_text.config(yscrollcommand=self.log_text_scrollbar.set)

        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.log_text_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # update current ui preset option checked
        self.current_preset.set(f"Preset {preset_number}")
        
        # update current ui checkbox for auto close
        if "auto_close" in self.config:
            self.auto_close_var.set(self.config.get("auto_close"))
        
        
        # Adding copyright and thanks message to the log output
        self.log(f"------------------------------------------------")
        self.log(f"Crypto90's WindowManager {current_version}. All rights reserved.")
        self.log(f"Thanks for using! Find the source code here:\nhttps://github.com/Crypto90/WindowManager")
        self.log(f"------------------------------------------------")
        self.log(f"Loaded preset number: {preset_number}")
        
        self.populate_window_list()
        
        self.stream_order()
        

        # If window states are loaded, restore main window's position and size
        if self.window_states:
            self.restore_main_window_position()
        
        
        
        
        self.process_listbox.bind("<Button-3>", self.show_context_menu)
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Add Launcher Override", command=self.set_launcher_override)

        
    def show_context_menu(self, event):
        try:
            widget = event.widget
            index = widget.nearest(event.y)
            text = widget.get(index)

            if text.startswith("★"):  # only allow context menu on starred entries
                self.process_listbox.selection_clear(0, tk.END)
                self.process_listbox.selection_set(index)
                self.context_menu.post(event.x_root, event.y_root)
        except Exception as e:
            print("Error: showing context menu:", e)
    
    def set_launcher_override(self):
        from tkinter import filedialog

        try:
            selection = self.process_listbox.curselection()
            if not selection:
                return
            index = selection[0]
            text = self.process_listbox.get(index)
            process_name = text.strip("★").split(" (")[0].strip()
            file_path = filedialog.askopenfilename(
                title="Select Launcher Executable",
                filetypes=[("Executable Files", "*.exe"), ("All Files", "*.*")]
            )
            if not file_path:
                return

            # Apply override directly to stored window states
            for key, state in self.window_states.items():
                # Match based on process_name
                if state.process_name == process_name:
                    state.launcher_override = file_path
                    self.log(f"Launcher override set for {key} -> {file_path}")
                    break  # Remove if multiple entries may match

            save_window_states(self.window_states, self.config)
            self.refresh_window_list()

        except Exception as e:
            self.log(f"Error: setting launcher override: {e}")

    #stdout=subprocess.DEVNULL,
    #stderr=subprocess.DEVNULL,
    #stdin=subprocess.DEVNULL,
    #creationflags=CREATE_NEW_CONSOLE if sys.platform == "win32" else 0,
    def launch_independent(self, path):
        CREATE_NEW_CONSOLE = 0x00000010
        try:
            subprocess.Popen(
                [path],
                cwd=os.path.dirname(path),
                creationflags=CREATE_NEW_CONSOLE if sys.platform == "win32" else 0,
                shell=False  # Only use True if path is a string command, not a real path
            )
        except Exception as e:
            print(f"Error: launching: {e}")
    
    def try_start_application(self, pname, state, window_states, save_path, auto_close_flag=False, UWP=False):
        
        launcher_path = state.launcher_override if state.launcher_override else state.process_path
        if os.path.isfile(launcher_path):
            # Check if the app is a UWP app by looking for "WindowsApps" in the path
            if "WindowsApps" in launcher_path:
                # Handle UWP app start using AppUserModelId or AppExecutionAlias
                self.log(f"Starting UWP app: {pname}")

                # Extract the app name from the path and remove the ".exe" extension if present
                app_name = launcher_path.split("\\")[-2]  # Extract the last part (the app name)
                #app_name_without_extension = app_name.replace(".exe", "")  # Remove .exe extension
                # Get the AppUserModelId using PowerShell
                uwp_app_name = get_uwp_app_name(app_name)
                
                if uwp_app_name:
                    # Start the UWP app using the AppUserModelId
                    uwp_start_command = f"cmd /c start {uwp_app_name}"
                    #print(uwp_start_command)
                    subprocess.Popen(uwp_start_command, shell=False)
                    
                    #self.launch_independent(uwp_start_command)
                    self.log(f"Started UWP app: {uwp_app_name}")
                else:
                    self.log(f"Error: Failed to find AppUserModelId for {pname}")
            else:
                self.launch_independent(launcher_path)
                self.log(f"Started app: {launcher_path}")
            return

        # Try to find the new path one level up
        base_dir = os.path.dirname(launcher_path)
        filename = os.path.basename(launcher_path)
        parent_dir = os.path.dirname(base_dir)

        for root, dirs, files in os.walk(parent_dir):
            if filename in files:
                new_path = os.path.join(root, filename)
                print(f"Info: App update with changed path detected: {new_path}")

                # Update paths in memory
                updated = False
                for state in window_states.values():
                    if launcher_path == app_path:
                        launcher_path = new_path
                        updated = True

                if updated:
                    config = {"auto_close": auto_close_flag}
                    save_window_states(window_states, config, filename=save_path)
                    print(f"Updated path saved to: {save_path}")

                
                if os.path.isfile(new_path):
                    # Check if the app is a UWP app by looking for "WindowsApps" in the path
                    if "WindowsApps" in new_path:
                        # Handle UWP app start using AppUserModelId or AppExecutionAlias
                        self.log(f"Starting UWP app: {pname}")

                        # Extract the app name from the path and remove the ".exe" extension if present
                        app_name = new_path.split("\\")[-2]  # Extract the last part (the app name)
                        #app_name_without_extension = app_name.replace(".exe", "")  # Remove .exe extension
                        # Get the AppUserModelId using PowerShell
                        uwp_app_name = get_uwp_app_name(app_name)
                        
                        #print("uwp app name:")
                        #print(uwp_app_name)

                        if uwp_app_name:
                            # Start the UWP app using the AppUserModelId
                            uwp_start_command = f"cmd /c start {uwp_app_name}"
                            #print(uwp_start_command)
                            subprocess.Popen(uwp_start_command, shell=True)
                            
                            #self.launch_independent(uwp_start_command)
                            self.log(f"Started UWP app: {uwp_app_name}")
                        else:
                            self.log(f"Failed to find AppUserModelId for {pname}")
                    else:
                        self.launch_independent(new_path)
                        self.log(f"Started app: {launcher_path}")
                    return
                return

        print(f"Error: Failed to find or start '{filename}' (original path: {app_path})")
    
    def switch_preset(self):
        selected = self.current_preset.get()
        # Determine the file corresponding to the current preset
        preset_file = f"window_states_{selected.split()[-1]}.pkl"
        
        self.window_states, self.config = load_window_states(preset_file)
        
        if "auto_close" in self.config:
            self.auto_close_var.set(self.config.get("auto_close"))
        else:
            self.auto_close_var.set(False)
        
        self.refresh_window_list()
        self.log(f"Switched to {selected}")

    
    def log(self, message):
        self.log_text.config(state=tk.NORMAL)

        # Define tags if they don't exist
        if "error" not in self.log_text.tag_names():
            self.log_text.tag_config("error", foreground="red")
        if "info" not in self.log_text.tag_names():
            self.log_text.tag_config("info", foreground="dark blue")

        # Determine which tag to use
        lower_msg = message.lower()
        if any(keyword in lower_msg for keyword in ["error", "exception", "fatal"]):
            tag = "error"
        elif "info" in lower_msg:
            tag = "info"
        else:
            tag = None

        # Insert the message with the appropriate tag
        if tag:
            self.log_text.insert(tk.END, message + "\n", tag)
        else:
            self.log_text.insert(tk.END, message + "\n")

        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)


    def refresh_window_list(self):
        self.process_listbox.delete(0, tk.END)
        self.window_mapping.clear()
        self.populate_window_list()

    def populate_window_list(self):
        current_names = {get_process_name_for_window(w) for w in gw.getAllWindows()}
        for pname, state in self.window_states.items():
            if pname == "main_window":
                continue  # Skip adding the main window
            if pname not in current_names:
                if state.launcher_override:
                    label = f"★★ {pname} (Not running)"
                else:
                    label = f"★ {pname} (Not running)"
                self.process_listbox.insert(tk.END, label)
                index = self.process_listbox.size() - 1
                self.process_listbox.itemconfig(index, {'fg': 'red'})
                self.window_mapping.append(None)
        
        grouped = get_visible_windows_grouped_by_monitor()
        for monitor, items in grouped.items():
            self.process_listbox.insert(tk.END, f"=== Monitor: {monitor} ===")
            self.process_listbox.itemconfig(tk.END, {'fg': '#2980b9'})
            self.window_mapping.append(None)
            for win, pname in items:
                saved = pname in self.window_states
                star = "★ " if saved else "  "
                
                saved_override = (
                    pname in self.window_states and
                    self.window_states[pname].launcher_override is not None
                )
                star = "★★ " if saved_override else star
                
                try:
                    window = gw.getWindowsWithTitle(win.title)[0]  # Get window with title
                    width = window.width
                    height = window.height
                    
                    hwnd = window._hWnd
                    placement = win32gui.GetWindowPlacement(hwnd)
                    hwnd_minimized = placement[1] == win32con.SW_SHOWMINIMIZED

                    if hwnd_minimized:
                        # Get restored position (normal state) from placement
                        restored_rect = placement[4]  # (left, top, right, bottom)
                        left = restored_rect[0]
                        top = restored_rect[1]
                        width = restored_rect[2] - restored_rect[0]
                        height = restored_rect[3] - restored_rect[1]
                    else:
                        left = window.left
                        top = window.top
                        width = window.width
                        height = window.height
                    
                    # Only skip if not minimized and size is 0x0
                    if (width > 0 and height > 0) or hwnd_minimized:
                        is_uwp = is_uwp_window(win)
                        uwp_tag = " [UWP]" if is_uwp else ""
                        label = f"{star}{pname} (Title: {win.title[:30]}, Size: {width}x{height}, Position: {left},{top}){uwp_tag}, Minimized: {hwnd_minimized}"
                    else:
                        continue  # Skip windows with size 0x0

                except Exception as e:
                    label = f"{star}{pname} (Title: {win.title[:30]}, Size: N/A, Position: N/A)"
                    print(f"Error: retrieving window size and position: {e}")

                self.process_listbox.insert(tk.END, label)
                index = self.process_listbox.size() - 1
                if saved:
                    state = self.window_states[pname]
                    running = is_process_running(state.process_path) if state.process_path else True
                    self.process_listbox.itemconfig(index, {'fg': 'green' if running else 'red'})
                    self.process_listbox.select_set(index)
                self.window_mapping.append((win, pname))

        


    def save_window_positions(self):
        selected_indices = self.process_listbox.curselection()
        self.window_states = {}
        
        selected = self.current_preset.get()
        preset_file = f"window_states_{selected.split()[-1]}.pkl"
        
        
        # Load existing states to preserve launcher_override and other data
        previous_states, _ = load_window_states(preset_file)

        for index in selected_indices:
            win_entry = self.window_mapping[index]

            if win_entry is None:
                continue  # Skip separators or non-window entries

            window, process_name = win_entry
            
            
            
            try:
                hwnd = window._hWnd
                hwnd_minimized = win32gui.IsIconic(hwnd)
                
                placement = win32gui.GetWindowPlacement(hwnd)
                hwnd_minimized = placement[1] == win32con.SW_SHOWMINIMIZED

                if hwnd_minimized:
                    restored_rect = placement[4]  # (left, top, right, bottom)
                    left = restored_rect[0]
                    top = restored_rect[1]
                    width = restored_rect[2] - restored_rect[0]
                    height = restored_rect[3] - restored_rect[1]
                else:
                    left = window.left
                    top = window.top
                    width = window.width
                    height = window.height

                position = (left, top)
                size = (width, height)

                
                path = get_process_path_for_window(window)
                uwp = is_uwp_window(window)
                url = None
                is_minimized = hwnd_minimized
                launcher_override_restore = None
                
                # Preserve launcher_override if it was previously set
                old_state = previous_states.get(process_name)
                if old_state and hasattr(old_state, "launcher_override"):
                    launcher_override_restore = old_state.launcher_override
                
                
                if uwp:
                    url = get_uwp_app_name(path)
                self.window_states[process_name] = WindowState(process_name, position, size, path, url, minimized=is_minimized, launcher_override=launcher_override_restore)
                self.log(f"Saved: {process_name} at {position} size {size}. Minimized: {is_minimized}")
            except Exception as e:
                self.log(f"Error: saving {process_name}: {e}")

        # Save config state
        config = {
            "auto_close": self.auto_close_var.get()
        }
        save_window_states(self.window_states, config, preset_file)
        self.log(f"Window positions saved for {selected}")
        self.refresh_window_list()
    
    def restore_main_window_position(self):
        if "main_window" in self.window_states:
            state = self.window_states["main_window"]
            self.root.geometry(f"{state.size[0]}x{state.size[1]}+{state.position[0]}+{state.position[1]}")
            self.log("Restored main window position and size.")


    def stream_order(self):
        
        # always refresh process list before we order
        self.refresh_window_list()
        
        started_any = False
        
        moved_resized_any = False
        
        
        selected = self.current_preset.get()
        preset_file = f"window_states_{selected.split()[-1]}.pkl"

        for pname, state in self.window_states.items():
            if not hasattr(state, "process_path") or state.process_path is None:
                continue  # Skip entries like "main_window" or invalid ones
            running = is_process_running(state.process_path)
            if not running and state.process_path:
                try:
                    self.try_start_application(pname, state, self.window_states, save_path=preset_file)
                    started_any = True
                except Exception as e:
                    self.log(f"Error: Failed to start '{pname} [{state.process_path}]': {e}")
        
        
        if started_any:
            self.root.update()
            self.log("Waiting for windows to appear...")

        max_wait_time = 30  # seconds
        waited = 0
        processed = set()

        while waited < max_wait_time:
            all_found = True

            for pname, state in self.window_states.items():
                if pname == "main_window" or pname in processed:
                    continue

                windows = [w for w in gw.getAllWindows()
                           if get_process_name_for_window(w) == pname and
                           (win32gui.IsWindowVisible(w._hWnd) or win32gui.IsIconic(w._hWnd))]

                if windows:
                    moved_resized_any = True
                    win = windows[0]
                    hwnd = win._hWnd
                    
                    try:
                        
                        # Check if window is maximized
                        placement = win32gui.GetWindowPlacement(hwnd)
                        if placement[1] == win32con.SW_SHOWMINIMIZED or placement[1] == win32con.SW_SHOWMAXIMIZED:
                            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                        
                        win.moveTo(*state.position)
                        win.resizeTo(*state.size)

                        if hasattr(state, 'minimized') and state.minimized:
                            win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
                    
                        self.log(f"Moved and resized window for {pname}")
                        processed.add(pname)
                        
                        
                    except PyGetWindowException as e:
                        if "Error code from Windows: 5" in str(e):
                            self.log(f"Error: Permission denied: Cannot move window for '{pname}'. Try running as administrator.")
                        else:
                            self.log(f"Error: Failed to move/resize window for '{pname}': {e}")
                    except Exception as e:
                        self.log(f"Error: Unexpected error while handling window '{pname}': {e}")

                else:
                    all_found = False

            if all_found:
                if moved_resized_any:
                    self.log(f"All processes moved and resized.")
                break

            time.sleep(1)
            waited += 1

        # Final check for any still-missing windows
        for pname in self.window_states:
            if pname != "main_window" and pname not in processed:
                self.log(f"Warning: Window for '{pname}' not found after waiting.")
        
        # Refresh the process list after the wait
        self.refresh_window_list()
        
        if self.auto_close_var.get():
            self.log("Auto-close is enabled. Exiting in 5 seconds...")
            self.root.after(5000, self.root.destroy)




def main():
    root = tk.Tk()
    root.geometry("600x500")
    app = WindowManagerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
