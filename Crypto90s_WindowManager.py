import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog
import psutil
import pickle
import json
import subprocess
import time
import os
import sys
import argparse
import threading
import datetime

# High-DPI and Windows-specific imports
IS_WINDOWS = sys.platform == "win32"

if IS_WINDOWS:
    import ctypes
    from ctypes import wintypes
    import win32process
    import win32gui
    import win32con
    import winreg
    from screeninfo import get_monitors
    import pygetwindow as gw
    from pygetwindow import PyGetWindowException
else:
    # Graceful stubs for cross-platform validation / development
    gw = None
    get_monitors = lambda: []
    PyGetWindowException = Exception

current_version = "v0.2.0"


def enable_high_dpi():
    """Enable Per-Monitor V2 DPI awareness on Windows to ensure accurate coordinates."""
    if not IS_WINDOWS:
        return
    try:
        # Per-Monitor V2 DPI awareness (Windows 10 1703+)
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            # Fallback to system DPI awareness (Windows Vista+)
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def get_data_dir():
    """Returns a reliable, writable directory for presets and config."""
    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    # Test if the directory is writable
    test_path = os.path.join(base_dir, ".wm_write_test")
    try:
        with open(test_path, "w") as f:
            f.write("ok")
        os.remove(test_path)
        return base_dir
    except (PermissionError, OSError):
        # Fallback to %LOCALAPPDATA%\Crypto90s_WindowManager
        appdata = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
        target_dir = os.path.join(appdata, "Crypto90s_WindowManager")
        os.makedirs(target_dir, exist_ok=True)
        return target_dir


def parse_args():
    parser = argparse.ArgumentParser(description="Crypto90's WindowManager Preset Selector")
    parser.add_argument("--preset", type=int, default=1, choices=range(1, 11),
                        help="Preset number to load (1-10, default: 1)")
    parser.add_argument("--silent", "--headless", action="store_true", dest="silent",
                        help="Run ordering in background without opening the GUI")
    parser.add_argument("--version", action="version", version=f"Crypto90's WindowManager {current_version}")
    return parser.parse_args()


# Class for window state
class WindowState:
    def __init__(self, process_name, position, size, process_path=None, url=None,
                 path=None, minimized=False, maximized=False, launcher_override=None,
                 window_title="", is_uwp=False, uwp_app_id=None):
        self.process_name = process_name
        self.position = list(position) if position else [0, 0]
        self.size = list(size) if size else [800, 600]
        self.process_path = process_path or path
        self.url = url
        self.path = self.process_path
        self.minimized = bool(minimized)
        self.maximized = bool(maximized)
        self.launcher_override = launcher_override
        self.window_title = window_title or ""
        self.is_uwp = bool(is_uwp)
        self.uwp_app_id = uwp_app_id

    def to_dict(self):
        return {
            "process_name": self.process_name,
            "window_title": self.window_title,
            "position": self.position,
            "size": self.size,
            "process_path": self.process_path,
            "minimized": self.minimized,
            "maximized": self.maximized,
            "launcher_override": self.launcher_override,
            "is_uwp": self.is_uwp,
            "uwp_app_id": self.uwp_app_id,
            "url": self.url
        }

    @classmethod
    def from_dict(cls, d):
        if not isinstance(d, dict):
            return None
        return cls(
            process_name=d.get("process_name", ""),
            position=d.get("position", [0, 0]),
            size=d.get("size", [800, 600]),
            process_path=d.get("process_path") or d.get("path"),
            url=d.get("url"),
            minimized=d.get("minimized", False),
            maximized=d.get("maximized", False),
            launcher_override=d.get("launcher_override"),
            window_title=d.get("window_title", ""),
            is_uwp=d.get("is_uwp", False),
            uwp_app_id=d.get("uwp_app_id")
        )


def get_preset_filepath(preset_number, ext=".json"):
    data_dir = get_data_dir()
    return os.path.join(data_dir, f"window_states_{preset_number}{ext}")


def load_window_states(preset_number=1):
    """
    Loads window states for a preset.
    Prefers JSON, with automatic migration from legacy .pkl files.
    """
    json_path = get_preset_filepath(preset_number, ".json")
    pkl_path = get_preset_filepath(preset_number, ".pkl")

    # 1. Try loading modern JSON format
    if os.path.isfile(json_path):
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                raw_states = data.get("window_states", {})
                config = data.get("config", {})
                window_states = {}
                for k, v in raw_states.items():
                    state_obj = WindowState.from_dict(v)
                    if state_obj:
                        window_states[k] = state_obj
                return window_states, config
        except Exception as e:
            print(f"Warning: Failed to load JSON preset {preset_number}: {e}")

    # 2. Fallback to legacy pickle format and auto-migrate
    if os.path.isfile(pkl_path):
        try:
            with open(pkl_path, "rb") as f:
                data = pickle.load(f)
            if isinstance(data, dict):
                raw_states = data.get("window_states", {})
                config = data.get("config", {})
                if not raw_states and not config:
                    raw_states = data  # Very old pickle format was just dict of states
                
                window_states = {}
                for k, v in raw_states.items():
                    if isinstance(v, WindowState):
                        window_states[k] = v
                    elif isinstance(v, dict):
                        state_obj = WindowState.from_dict(v)
                        if state_obj:
                            window_states[k] = state_obj
                    elif hasattr(v, "__dict__"):
                        state_obj = WindowState.from_dict(v.__dict__)
                        if state_obj:
                            window_states[k] = state_obj

                # Auto-migrate to JSON
                save_window_states(window_states, config, preset_number)
                print(f"Info: Migrated preset {preset_number} from PKL to JSON format.")
                return window_states, config
        except Exception as e:
            print(f"Warning: Failed to load legacy PKL preset {preset_number}: {e}")

    return {}, {}


def save_window_states(window_states, config=None, preset_number=1):
    """Saves window states and configuration to a clean JSON file."""
    json_path = get_preset_filepath(preset_number, ".json")
    serialized_states = {}
    for k, v in window_states.items():
        if isinstance(v, WindowState):
            serialized_states[k] = v.to_dict()
        elif isinstance(v, dict):
            serialized_states[k] = v

    data = {
        "version": current_version,
        "window_states": serialized_states,
        "config": config or {}
    }

    try:
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except Exception as e:
        print(f"Error: Failed to save preset to {json_path}: {e}")


def get_monitor_for_window(win, monitors):
    if not IS_WINDOWS or not monitors:
        return None
    hwnd = win._hWnd
    try:
        if win32gui.IsIconic(hwnd):
            placement = win32gui.GetWindowPlacement(hwnd)
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

    return monitors[0] if monitors else None


def get_process_info_for_window(win):
    """Returns (process_name, process_path, pid) for a window handle."""
    if not IS_WINDOWS:
        return None, None, None
    try:
        _, pid = win32process.GetWindowThreadProcessId(win._hWnd)
        proc = psutil.Process(pid)
        return proc.name(), proc.exe(), pid
    except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess, Exception):
        return None, None, None


def get_running_process_paths():
    """Fast snapshot of all running executable paths."""
    paths = set()
    for proc in psutil.process_iter(['exe']):
        try:
            exe = proc.info.get('exe')
            if exe:
                paths.add(os.path.normcase(exe))
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            continue
    return paths


def is_process_running(process_path, running_cache=None):
    if not process_path:
        return False
    norm_path = os.path.normcase(process_path)
    if running_cache is not None:
        return norm_path in running_cache

    for proc in psutil.process_iter(['exe']):
        try:
            exe = proc.info.get('exe')
            if exe and os.path.normcase(exe) == norm_path:
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            continue
    return False


def get_visible_windows_grouped_by_monitor():
    """
    Groups open application windows by monitor.
    Filters out shell/desktop windows and WindowManager itself, while preserving File Explorer.
    """
    if not IS_WINDOWS:
        return {}

    grouped = {}
    monitors = get_monitors()
    current_pid = os.getpid()

    for window in gw.getAllWindows():
        hWnd = window._hWnd
        process_name, process_path, pid = get_process_info_for_window(window)

        if not process_name:
            continue

        # Exclude WindowManager itself
        if pid == current_pid:
            continue

        # For explorer.exe, allow CabinetWClass (actual folder windows), ignore Desktop/Taskbar
        if process_name.lower() == "explorer.exe":
            class_name = win32gui.GetClassName(hWnd)
            if class_name != "CabinetWClass":
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
            mon_id = f"{monitor.x}x{monitor.y} ({monitor.width}x{monitor.height})"
        else:
            mon_id = "Default Display"

        if mon_id not in grouped:
            grouped[mon_id] = []
        grouped[mon_id].append((window, process_name, process_path))

    return grouped


def is_uwp_window(window):
    if not IS_WINDOWS:
        return False
    _, path, _ = get_process_info_for_window(window)
    return bool(path and "WindowsApps" in path)


def get_uwp_app_info(path):
    """
    Extracts package folder name and looks up AppUserModelId (AUMID) for UWP applications.
    """
    if not path or "WindowsApps" not in path:
        return None, None
    try:
        parts = path.split("\\")
        for part in parts:
            if "__" in part:
                package_folder = part
                identifier = package_folder.split("__")[-1]
                # Query PowerShell for AppID
                cmd = ["powershell", "-NoProfile", "-Command", "Get-StartApps | Format-List Name, AppID"]
                proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=False)
                stdout, _ = proc.communicate(timeout=5)
                lines = stdout.decode("utf-8", errors="ignore").splitlines()
                
                name = None
                appid = None
                for line in lines:
                    if line.startswith("Name"):
                        name = line.split(":", 1)[1].strip()
                    elif line.startswith("AppID"):
                        curr_appid = line.split(":", 1)[1].strip()
                        if identifier.lower() in curr_appid.lower():
                            appid = curr_appid
                            return name, appid
                break
    except Exception as e:
        print(f"Error resolving UWP info: {e}")
    return None, None


def is_start_with_windows_enabled():
    """Checks if WindowManager is configured to run at Windows startup."""
    if not IS_WINDOWS:
        return False
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_READ)
        val, _ = winreg.QueryValueEx(key, "Crypto90s_WindowManager")
        winreg.CloseKey(key)
        return bool(val)
    except FileNotFoundError:
        return False
    except Exception:
        return False


def set_start_with_windows(enabled, preset=1):
    """Enables or disables WindowManager launch at Windows startup."""
    if not IS_WINDOWS:
        return False
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_SET_VALUE)
        if enabled:
            if getattr(sys, "frozen", False):
                exe_cmd = f'"{sys.executable}" --preset {preset} --silent'
            else:
                exe_cmd = f'"{sys.executable}" "{os.path.abspath(__file__)}" --preset {preset} --silent'
            winreg.SetValueEx(key, "Crypto90s_WindowManager", 0, winreg.REG_SZ, exe_cmd)
        else:
            try:
                winreg.DeleteValue(key, "Crypto90s_WindowManager")
            except FileNotFoundError:
                pass
        winreg.CloseKey(key)
        return True
    except Exception as e:
        print(f"Error setting Windows startup registry: {e}")
        return False


class WindowManagerApp:
    def __init__(self, root, initial_preset=1):
        self.root = root
        self.root.title("Crypto90's WindowManager")
        self.current_preset_num = initial_preset

        self.window_states, self.config = load_window_states(self.current_preset_num)
        self.window_mapping = []  # Stores mapping for listbox rows: (key, window_obj, pname, ppath) or None
        self.ordering_in_progress = False
        self.cancel_ordering = False
        self.auto_close_job = None

        self.root.minsize(width=580, height=480)
        self.root.configure(bg="#1e1e1e")

        # Top Bar: Process listbox with scrollbar
        self.process_listbox_frame = tk.Frame(self.root, bg="#1e1e1e")
        self.process_listbox_frame.pack(padx=10, pady=8, fill=tk.BOTH, expand=True)

        self.process_listbox = tk.Listbox(
            self.process_listbox_frame,
            selectmode=tk.MULTIPLE,
            bg="#252525",
            fg="#f0f0f0",
            selectbackground="#1e6b37",
            selectforeground="#ffffff",
            highlightbackground="#3d3d3d",
            font=("Segoe UI" if IS_WINDOWS else "Helvetica", 9),
            relief=tk.FLAT
        )
        self.process_listbox_scrollbar = tk.Scrollbar(
            self.process_listbox_frame,
            orient="vertical",
            command=self.process_listbox.yview,
            troughcolor="#252525",
            bg="#444"
        )
        self.process_listbox.config(yscrollcommand=self.process_listbox_scrollbar.set)
        self.process_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.process_listbox_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Preset Management Frame
        preset_frame = tk.Frame(self.root, bg="#1e1e1e")
        preset_frame.pack(pady=4)

        self.current_preset = tk.StringVar(value=f"Preset {self.current_preset_num}")
        tk.Label(preset_frame, text="Presets:", fg="#ffffff", bg="#1e1e1e", font=("Segoe UI", 9, "bold")).pack(side="left", padx=5)

        self.preset_radios = []
        for i in range(1, 11):
            name = f"Preset {i}"
            rb = tk.Radiobutton(
                preset_frame, text=str(i), variable=self.current_preset, value=name,
                command=self.switch_preset, bg="#1e1e1e", fg="#ffffff", selectcolor="#333333",
                activebackground="#1e1e1e", activeforeground="#3498db"
            )
            rb.pack(side="left", padx=2)
            self.preset_radios.append(rb)

        tk.Button(preset_frame, text="✏️ Rename", bg="#3d3d3d", fg="#ffffff", relief=tk.FLAT,
                  padx=5, font=("Segoe UI", 8), command=self.rename_current_preset).pack(side="left", padx=6)

        # Options Frame: Auto-close & Run at Windows Startup
        options_frame = tk.Frame(self.root, bg="#1e1e1e")
        options_frame.pack(pady=4)

        self.auto_close_var = tk.BooleanVar(value=self.config.get("auto_close", False))
        self.auto_close_checkbox = tk.Checkbutton(
            options_frame, text="Auto-close after ordering", variable=self.auto_close_var,
            command=self.on_auto_close_toggle, bg="#1e1e1e", fg="#ffffff", selectcolor="#333333",
            activebackground="#1e1e1e", activeforeground="#ffffff"
        )
        self.auto_close_checkbox.pack(side="left", padx=10)

        self.startup_var = tk.BooleanVar(value=is_start_with_windows_enabled())
        self.startup_checkbox = tk.Checkbutton(
            options_frame, text="Run at Windows Startup", variable=self.startup_var,
            command=self.on_startup_toggle, bg="#1e1e1e", fg="#ffffff", selectcolor="#333333",
            activebackground="#1e1e1e", activeforeground="#ffffff"
        )
        self.startup_checkbox.pack(side="left", padx=10)

        # Action Button Frame
        button_frame = tk.Frame(self.root, bg="#1e1e1e")
        button_frame.pack(pady=6)

        tk.Button(button_frame, text="🔄 Refresh Windows", bg="#2980b9", fg="white", relief=tk.FLAT,
                  command=self.refresh_window_list, padx=8, pady=3).pack(side="left", padx=4)
        tk.Button(button_frame, text="💾 Save Selected Positions", bg="#7f8c8d", fg="white", relief=tk.FLAT,
                  command=self.save_window_positions, padx=8, pady=3).pack(side="left", padx=4)

        self.order_button = tk.Button(button_frame, text="🚀 Start, Resize & Order", bg="#27ae60", fg="white",
                                      relief=tk.FLAT, command=self.toggle_stream_order, padx=10, pady=3,
                                      font=("Segoe UI", 9, "bold"))
        self.order_button.pack(side="left", padx=4)

        tk.Button(button_frame, text="☕ Buy Coffee", bg="#e67e22", fg="white", relief=tk.FLAT,
                  command=lambda: webbrowser.open("https://ko-fi.com/crypto90"), padx=6, pady=3).pack(side="left", padx=4)

        # Log box frame (Packed at bottom)
        self.log_text_frame = tk.Frame(self.root, bg="#1e1e1e")
        self.log_text_frame.pack(padx=10, pady=6, fill=tk.X, side=tk.BOTTOM)

        self.log_text = tk.Text(
            self.log_text_frame,
            height=8,
            state=tk.DISABLED,
            bg="#181818",
            fg="#e0e0e0",
            insertbackground="white",
            highlightbackground="#333",
            font=("Consolas" if IS_WINDOWS else "Courier", 9),
            relief=tk.FLAT
        )

        self.log_text_scrollbar = tk.Scrollbar(
            self.log_text_frame,
            orient="vertical",
            command=self.log_text.yview,
            troughcolor="#181818",
            bg="#333"
        )
        self.log_text.config(yscrollcommand=self.log_text_scrollbar.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.log_text_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Pre-configure log color tags (Fixes dark-blue unreadability bug)
        self.log_text.tag_config("error", foreground="#e74c3c")    # Bright Red
        self.log_text.tag_config("info", foreground="#3498db")     # Vibrant Cyan / Blue
        self.log_text.tag_config("success", foreground="#2ecc71")  # Bright Green
        self.log_text.tag_config("warn", foreground="#f39c12")     # Amber / Orange
        self.log_text.tag_config("time", foreground="#888888")     # Muted timestamp

        # Context Menu
        self.context_menu = tk.Menu(self.root, tearoff=0, bg="#2d2d2d", fg="white", activebackground="#3498db")
        self.context_menu.add_command(label="Add Launcher Override...", command=self.set_launcher_override)
        self.context_menu.add_command(label="Remove Launcher Override", command=self.remove_launcher_override)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Open File Location", command=self.open_file_location)
        self.process_listbox.bind("<Button-3>", self.show_context_menu)
        if sys.platform == "darwin":
            self.process_listbox.bind("<Button-2>", self.show_context_menu)

        # Initial Welcome Log
        self.log("------------------------------------------------", tag="info")
        self.log(f"Crypto90's WindowManager {current_version} Initialized", tag="success")
        self.log(f"Storage: {get_data_dir()}", tag="info")
        self.log("------------------------------------------------", tag="info")

        self.update_preset_labels()
        self.populate_window_list()

        # Restore main window geometry if saved
        if "main_window" in self.window_states:
            self.restore_main_window_position()

        # Execute initial order safely after GUI has mounted
        self.root.after(300, self.start_stream_order)

    def log(self, message, tag=None):
        """Thread-safe, timestamped logging to the UI log window."""
        def _append():
            self.log_text.config(state=tk.NORMAL)
            timestamp = datetime.datetime.now().strftime("[%H:%M:%S] ")
            self.log_text.insert(tk.END, timestamp, "time")

            chosen_tag = tag
            if not chosen_tag:
                lower = message.lower()
                if any(w in lower for keyword in ["error", "exception", "failed", "denied"] for w in [keyword]):
                    chosen_tag = "error"
                elif any(w in lower for keyword in ["success", "saved", "restored", "started"] for w in [keyword]):
                    chosen_tag = "success"
                elif "warn" in lower:
                    chosen_tag = "warn"
                elif "info" in lower:
                    chosen_tag = "info"

            if chosen_tag:
                self.log_text.insert(tk.END, message + "\n", chosen_tag)
            else:
                self.log_text.insert(tk.END, message + "\n")

            self.log_text.see(tk.END)
            self.log_text.config(state=tk.DISABLED)

        if threading.current_thread() is threading.main_thread():
            _append()
        else:
            self.root.after(0, _append)

    def on_auto_close_toggle(self):
        self.config["auto_close"] = self.auto_close_var.get()
        save_window_states(self.window_states, self.config, self.current_preset_num)
        self.log(f"Auto-close set to {self.auto_close_var.get()}")

    def on_startup_toggle(self):
        enabled = self.startup_var.get()
        success = set_start_with_windows(enabled, self.current_preset_num)
        if success:
            self.log(f"Windows startup launch {'enabled' if enabled else 'disabled'}.", tag="success")
        else:
            self.log("Failed to update Windows startup registry.", tag="error")
            self.startup_var.set(not enabled)

    def update_preset_labels(self):
        preset_names = self.config.get("preset_names", {})
        for idx, rb in enumerate(self.preset_radios, start=1):
            custom_name = preset_names.get(str(idx))
            if custom_name:
                rb.config(text=f"{idx}: {custom_name}")
            else:
                rb.config(text=str(idx))

    def rename_current_preset(self):
        preset_names = self.config.get("preset_names", {})
        curr = preset_names.get(str(self.current_preset_num), f"Preset {self.current_preset_num}")
        new_name = simpledialog.askstring(
            "Rename Preset",
            f"Enter custom name for Preset {self.current_preset_num}:",
            initialvalue=curr,
            parent=self.root
        )
        if new_name is not None:
            new_name = new_name.strip()
            if new_name:
                preset_names[str(self.current_preset_num)] = new_name
            else:
                preset_names.pop(str(self.current_preset_num), None)
            self.config["preset_names"] = preset_names
            save_window_states(self.window_states, self.config, self.current_preset_num)
            self.update_preset_labels()
            self.log(f"Preset {self.current_preset_num} renamed to '{new_name or 'Default'}'.")

    def switch_preset(self):
        selected_text = self.current_preset.get()
        num = int(selected_text.split()[1])
        self.current_preset_num = num

        self.window_states, self.config = load_window_states(self.current_preset_num)
        self.auto_close_var.set(self.config.get("auto_close", False))
        self.update_preset_labels()
        self.refresh_window_list()
        self.log(f"Switched to Preset {self.current_preset_num}")

    def show_context_menu(self, event):
        try:
            index = self.process_listbox.nearest(event.y)
            if index < 0 or index >= len(self.window_mapping):
                return
            entry = self.window_mapping[index]
            if not entry:
                return  # Header row

            self.process_listbox.selection_clear(0, tk.END)
            self.process_listbox.selection_set(index)
            self.context_menu.post(event.x_root, event.y_root)
        except Exception as e:
            print(f"Error showing context menu: {e}")

    def get_selected_mapping_entry(self):
        selection = self.process_listbox.curselection()
        if not selection:
            return None
        index = selection[0]
        if index < len(self.window_mapping):
            return self.window_mapping[index]
        return None

    def set_launcher_override(self):
        entry = self.get_selected_mapping_entry()
        if not entry:
            return
        state_key, _, pname, _ = entry

        file_path = filedialog.askopenfilename(
            title=f"Select Launcher Executable for {pname}",
            filetypes=[("Executable Files", "*.exe"), ("All Files", "*.*")]
        )
        if not file_path:
            return

        if state_key in self.window_states:
            self.window_states[state_key].launcher_override = file_path
        else:
            # Create a placeholder state if not saved yet
            self.window_states[state_key] = WindowState(
                process_name=pname, position=[100, 100], size=[800, 600],
                process_path=file_path, launcher_override=file_path
            )

        save_window_states(self.window_states, self.config, self.current_preset_num)
        self.log(f"Launcher override for {pname} -> {file_path}", tag="success")
        self.refresh_window_list()

    def remove_launcher_override(self):
        entry = self.get_selected_mapping_entry()
        if not entry:
            return
        state_key, _, pname, _ = entry

        if state_key in self.window_states and self.window_states[state_key].launcher_override:
            self.window_states[state_key].launcher_override = None
            save_window_states(self.window_states, self.config, self.current_preset_num)
            self.log(f"Removed launcher override for {pname}", tag="info")
            self.refresh_window_list()

    def open_file_location(self):
        entry = self.get_selected_mapping_entry()
        if not entry:
            return
        state_key, _, pname, ppath = entry

        path = None
        if state_key in self.window_states:
            st = self.window_states[state_key]
            path = st.launcher_override or st.process_path
        if not path:
            path = ppath

        if path and os.path.exists(path):
            folder = os.path.dirname(path)
            if IS_WINDOWS:
                subprocess.Popen(f'explorer.exe /select,"{path}"')
            else:
                subprocess.Popen(["open", folder])
        else:
            self.log(f"Executable path not found on disk: {path}", tag="warn")

    def refresh_window_list(self):
        self.process_listbox.delete(0, tk.END)
        self.window_mapping.clear()
        self.populate_window_list()

    def populate_window_list(self):
        running_cache = get_running_process_paths()
        grouped = get_visible_windows_grouped_by_monitor()

        # Track keys currently running on screen
        current_running_keys = set()
        for items in grouped.values():
            for win, pname, ppath in items:
                title = win.title if win else ""
                key = f"{pname}::{title[:40]}" if title else pname
                current_running_keys.add(key)
                current_running_keys.add(pname)

        # 1. Render Saved Applications that are NOT currently running
        for key, state in self.window_states.items():
            if key == "main_window":
                continue
            if state.process_name not in current_running_keys and key not in current_running_keys:
                star = "★★ " if state.launcher_override else "★ "
                title_suffix = f" [{state.window_title[:25]}]" if state.window_title else ""
                label = f"{star}{state.process_name}{title_suffix} (Not running)"
                self.process_listbox.insert(tk.END, label)
                idx = self.process_listbox.size() - 1
                self.process_listbox.itemconfig(idx, {'fg': '#e74c3c'})  # Red
                # Store entry so offline apps are never deleted on save
                self.window_mapping.append((key, None, state.process_name, state.process_path))

        # 2. Render Currently Open Windows grouped by Monitor
        for monitor_name, items in grouped.items():
            self.process_listbox.insert(tk.END, f"=== Monitor: {monitor_name} ===")
            self.process_listbox.itemconfig(tk.END, {'fg': '#3498db'})
            self.window_mapping.append(None)

            for win, pname, ppath in items:
                title = win.title if win else ""
                key = f"{pname}::{title[:40]}" if title else pname

                saved = key in self.window_states or pname in self.window_states
                active_state = self.window_states.get(key) or self.window_states.get(pname)
                saved_override = bool(active_state and active_state.launcher_override)

                star = "★★ " if saved_override else ("★ " if saved else "  ")

                try:
                    hwnd = win._hWnd
                    placement = win32gui.GetWindowPlacement(hwnd) if IS_WINDOWS else (0, 1, 0, 0, (0, 0, 800, 600))
                    hwnd_minimized = placement[1] == win32con.SW_SHOWMINIMIZED if IS_WINDOWS else False
                    hwnd_maximized = placement[1] == win32con.SW_SHOWMAXIMIZED if IS_WINDOWS else False

                    if hwnd_minimized or hwnd_maximized:
                        restored_rect = placement[4]
                        left = restored_rect[0]
                        top = restored_rect[1]
                        width = restored_rect[2] - restored_rect[0]
                        height = restored_rect[3] - restored_rect[1]
                    else:
                        left = win.left
                        top = win.top
                        width = win.width
                        height = win.height

                    uwp_tag = " [UWP]" if is_uwp_window(win) else ""
                    state_tag = ", Minimized" if hwnd_minimized else (", Maximized" if hwnd_maximized else "")
                    label = f"{star}{pname} (Title: {title[:28]}, Size: {width}x{height}, Pos: {left},{top}){uwp_tag}{state_tag}"
                except Exception as e:
                    label = f"{star}{pname} (Title: {title[:28]}, Size: N/A)"

                self.process_listbox.insert(tk.END, label)
                idx = self.process_listbox.size() - 1

                if saved:
                    running = is_process_running(ppath or active_state.process_path, running_cache)
                    self.process_listbox.itemconfig(idx, {'fg': '#2ecc71' if running else '#e74c3c'})
                    self.process_listbox.select_set(idx)

                self.window_mapping.append((key, win, pname, ppath))

    def save_window_positions(self):
        """
        Saves selected window positions.
        CRITICAL BUG FIX: Preserves offline/closed applications in the preset!
        """
        selected_indices = self.process_listbox.curselection()
        if not selected_indices:
            self.log("No windows selected to save.", tag="warn")
            return

        # Load previous states to preserve offline apps and overrides
        previous_states, _ = load_window_states(self.current_preset_num)
        new_window_states = {}

        # Preserve any offline windows that were previously saved and are selected
        for idx in selected_indices:
            if idx >= len(self.window_mapping):
                continue
            entry = self.window_mapping[idx]
            if not entry:
                continue  # Skip monitor headers

            state_key, win, pname, ppath = entry

            # Case A: Offline / non-running window was selected -> Preserve from previous states
            if win is None:
                if state_key in previous_states:
                    new_window_states[state_key] = previous_states[state_key]
                    self.log(f"Preserved offline app: {pname}", tag="info")
                continue

            # Case B: Currently open window
            try:
                hwnd = win._hWnd
                placement = win32gui.GetWindowPlacement(hwnd) if IS_WINDOWS else (0, 1, 0, 0, (0, 0, 800, 600))
                is_minimized = placement[1] == win32con.SW_SHOWMINIMIZED if IS_WINDOWS else False
                is_maximized = placement[1] == win32con.SW_SHOWMAXIMIZED if IS_WINDOWS else False

                if is_minimized or is_maximized:
                    restored = placement[4]
                    left = restored[0]
                    top = restored[1]
                    width = restored[2] - restored[0]
                    height = restored[3] - restored[1]
                else:
                    left = win.left
                    top = win.top
                    width = win.width
                    height = win.height

                # Preserve launcher override if previously defined
                launcher_override = None
                if state_key in previous_states:
                    launcher_override = previous_states[state_key].launcher_override
                elif pname in previous_states:
                    launcher_override = previous_states[pname].launcher_override

                is_uwp = is_uwp_window(win)
                uwp_app_name, uwp_app_id = (None, None)
                if is_uwp and ppath:
                    uwp_app_name, uwp_app_id = get_uwp_app_info(ppath)

                st = WindowState(
                    process_name=pname,
                    position=(left, top),
                    size=(width, height),
                    process_path=ppath,
                    minimized=is_minimized,
                    maximized=is_maximized,
                    launcher_override=launcher_override,
                    window_title=win.title,
                    is_uwp=is_uwp,
                    uwp_app_id=uwp_app_id
                )
                new_window_states[state_key] = st
                self.log(f"Saved: {pname} ({width}x{height} at {left},{top})", tag="success")
            except Exception as e:
                self.log(f"Error saving window {pname}: {e}", tag="error")

        self.window_states = new_window_states
        save_window_states(self.window_states, self.config, self.current_preset_num)
        self.log(f"Saved {len(self.window_states)} window state(s) to Preset {self.current_preset_num}.", tag="success")
        self.refresh_window_list()

    def restore_main_window_position(self):
        if "main_window" in self.window_states:
            state = self.window_states["main_window"]
            self.root.geometry(f"{state.size[0]}x{state.size[1]}+{state.position[0]}+{state.position[1]}")
            self.log("Restored main window position and size.", tag="info")

    def toggle_stream_order(self):
        if self.ordering_in_progress:
            self.cancel_ordering = True
            self.log("Cancelling ordering operation...", tag="warn")
        else:
            self.start_stream_order()

    def start_stream_order(self):
        """Starts window ordering in a non-blocking background thread."""
        if self.ordering_in_progress:
            return
        self.ordering_in_progress = True
        self.cancel_ordering = False
        self.order_button.config(text="⏹ Stop / Cancel", bg="#c0392b")

        # Cancel any pending auto-close timer if re-running
        if self.auto_close_job:
            self.root.after_cancel(self.auto_close_job)
            self.auto_close_job = None

        thread = threading.Thread(target=self._stream_order_worker, daemon=True)
        thread.start()

    def _stream_order_worker(self):
        """Background worker thread: launches missing apps and repositions windows."""
        try:
            self.root.after(0, self.refresh_window_list)
            time.sleep(0.3)

            running_cache = get_running_process_paths()
            started_any = False

            # Phase 1: Launch any processes that are not currently running
            for key, state in list(self.window_states.items()):
                if self.cancel_ordering:
                    break
                if key == "main_window" or not state.process_path:
                    continue

                if not is_process_running(state.process_path, running_cache):
                    try:
                        self.try_start_application(state)
                        started_any = True
                    except Exception as e:
                        self.log(f"Failed to start '{state.process_name}': {e}", tag="error")

            if started_any:
                self.log("Waiting for started applications to initialize...", tag="info")
                time.sleep(2.0)

            # Phase 2: Poll and reposition windows
            max_wait = 25
            waited = 0
            processed_keys = set()

            while waited < max_wait and not self.cancel_ordering:
                all_found = True
                if not IS_WINDOWS:
                    break

                for key, state in list(self.window_states.items()):
                    if key == "main_window" or key in processed_keys or self.cancel_ordering:
                        continue

                    # Search for matching window
                    matched_win = None
                    for w in gw.getAllWindows():
                        pname, ppath, _ = get_process_info_for_window(w)
                        if pname == state.process_name:
                            # If title is saved, try to match title; otherwise match process
                            if state.window_title and state.window_title in w.title:
                                matched_win = w
                                break
                            elif not matched_win:
                                matched_win = w

                    if matched_win:
                        hwnd = matched_win._hWnd
                        try:
                            # Unmaximize/restore before moving
                            placement = win32gui.GetWindowPlacement(hwnd)
                            if placement[1] in (win32con.SW_SHOWMINIMIZED, win32con.SW_SHOWMAXIMIZED):
                                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

                            matched_win.moveTo(*state.position)
                            matched_win.resizeTo(*state.size)

                            # Re-apply maximized or minimized state
                            if state.maximized:
                                win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                            elif state.minimized:
                                win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)

                            self.log(f"Arranged: {state.process_name}", tag="success")
                            processed_keys.add(key)
                        except PyGetWindowException as e:
                            if "Error code from Windows: 5" in str(e):
                                self.log(f"Permission denied for '{state.process_name}'. Run as Administrator.", tag="error")
                            else:
                                self.log(f"Failed to position '{state.process_name}': {e}", tag="error")
                        except Exception as e:
                            self.log(f"Error handling '{state.process_name}': {e}", tag="error")
                    else:
                        all_found = False

                if all_found or len(processed_keys) == len([k for k in self.window_states if k != "main_window"]):
                    self.log("All configured application windows arranged!", tag="success")
                    break

                time.sleep(1.0)
                waited += 1

            if self.cancel_ordering:
                self.log("Ordering stopped by user.", tag="warn")

        except Exception as e:
            self.log(f"Unexpected error in ordering worker: {e}", tag="error")
        finally:
            self.ordering_in_progress = False
            self.root.after(0, self._on_ordering_finished)

    def _on_ordering_finished(self):
        self.order_button.config(text="🚀 Start, Resize & Order", bg="#27ae60")
        self.refresh_window_list()

        # Handle auto-close
        if self.auto_close_var.get() and not self.cancel_ordering:
            self.log("Auto-close is enabled. Exiting in 5 seconds... (Click 'Stop' to cancel)", tag="info")
            self.order_button.config(text="Cancel Auto-Close", bg="#e67e22", command=self.cancel_auto_close)
            self.auto_close_job = self.root.after(5000, self.root.destroy)

    def cancel_auto_close(self):
        if self.auto_close_job:
            self.root.after_cancel(self.auto_close_job)
            self.auto_close_job = None
            self.log("Auto-close cancelled.", tag="info")
            self.order_button.config(text="🚀 Start, Resize & Order", bg="#27ae60", command=self.toggle_stream_order)

    def launch_independent(self, path):
        """Launches an independent process cleanly."""
        CREATE_NEW_CONSOLE = 0x00000010
        try:
            cwd = os.path.dirname(path) if os.path.isfile(path) else None
            subprocess.Popen(
                [path],
                cwd=cwd,
                creationflags=CREATE_NEW_CONSOLE if IS_WINDOWS else 0,
                shell=False
            )
        except OSError as e:
            if hasattr(e, 'winerror') and e.winerror == 740:  # Elevation required
                self.log(f"Elevation required for {os.path.basename(path)}. Launching via ShellExecute...", tag="warn")
                if IS_WINDOWS:
                    ctypes.windll.shell32.ShellExecuteW(None, "open", path, None, None, 1)
            else:
                raise

    def try_start_application(self, state):
        """
        Starts an application.
        CRITICAL BUG FIX: Fixes undefined 'app_path', stops runaway root walks, handles UWP correctly.
        """
        pname = state.process_name
        launcher_path = state.launcher_override or state.process_path
        if not launcher_path:
            self.log(f"No executable path stored for {pname}", tag="warn")
            return

        # 1. Handle UWP Applications
        if state.is_uwp or "WindowsApps" in launcher_path:
            uwp_id = state.uwp_app_id
            if not uwp_id:
                _, uwp_id = get_uwp_app_info(launcher_path)
                state.uwp_app_id = uwp_id

            if uwp_id:
                self.log(f"Starting UWP app: {pname} ({uwp_id})", tag="info")
                subprocess.Popen(f'explorer.exe shell:AppsFolder\\{uwp_id}', shell=True)
                return
            else:
                self.log(f"Starting UWP app via shell fallback: {pname}", tag="info")
                subprocess.Popen(f'cmd /c start "" "{launcher_path}"', shell=True)
                return

        # 2. Regular Win32 Path Exists
        if os.path.isfile(launcher_path):
            self.launch_independent(launcher_path)
            self.log(f"Started: {launcher_path}", tag="success")
            return

        # 3. Fallback: Search for relocated executable (e.g. app update)
        # SAFEGUARD: Never walk root drive C:\ or deep directory trees
        base_dir = os.path.dirname(launcher_path)
        filename = os.path.basename(launcher_path)
        parent_dir = os.path.dirname(base_dir)

        # Ensure parent_dir is not root or empty (e.g., 'C:\\')
        if parent_dir and len(parent_dir.strip(":/\\")) > 1 and os.path.isdir(parent_dir):
            for root, dirs, files in os.walk(parent_dir):
                # Restrict search depth to 2 levels
                rel = os.path.relpath(root, parent_dir)
                if rel.count(os.sep) > 2:
                    continue
                if filename in files:
                    new_path = os.path.join(root, filename)
                    if os.path.isfile(new_path):
                        self.log(f"Detected updated path for {pname}: {new_path}", tag="info")
                        # Update state in memory
                        state.process_path = new_path
                        save_window_states(self.window_states, self.config, self.current_preset_num)
                        self.launch_independent(new_path)
                        return

        self.log(f"Could not find or launch: {launcher_path}", tag="error")


def run_headless(preset_number=1):
    """Headless CLI mode: orders windows silently in background without GUI."""
    print(f"Crypto90's WindowManager {current_version} [Headless Mode]")
    states, config = load_window_states(preset_number)
    if not states:
        print(f"No window states found for Preset {preset_number}. Exiting.")
        return

    running_cache = get_running_process_paths()
    for state in states.values():
        launcher = state.launcher_override or state.process_path
        if launcher and not is_process_running(launcher, running_cache):
            try:
                if state.is_uwp and state.uwp_app_id:
                    subprocess.Popen(f'explorer.exe shell:AppsFolder\\{state.uwp_app_id}', shell=True)
                elif os.path.isfile(launcher):
                    subprocess.Popen([launcher], cwd=os.path.dirname(launcher), shell=False)
            except Exception as e:
                print(f"Failed to start {state.process_name}: {e}")

    time.sleep(2.0)
    if not IS_WINDOWS:
        return

    waited = 0
    while waited < 20:
        all_done = True
        for state in states.values():
            for w in gw.getAllWindows():
                pname, _, _ = get_process_info_for_window(w)
                if pname == state.process_name:
                    hwnd = w._hWnd
                    placement = win32gui.GetWindowPlacement(hwnd)
                    if placement[1] in (win32con.SW_SHOWMINIMIZED, win32con.SW_SHOWMAXIMIZED):
                        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    w.moveTo(*state.position)
                    w.resizeTo(*state.size)
                    if state.maximized:
                        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                    elif state.minimized:
                        win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
                    break
            else:
                all_done = False
        if all_done:
            break
        time.sleep(1.0)
        waited += 1
    print("Headless window ordering complete.")


def main():
    enable_high_dpi()
    args = parse_args()

    if args.silent:
        run_headless(args.preset)
        return

    root = tk.Tk()
    root.geometry("640x520")
    app = WindowManagerApp(root, initial_preset=args.preset)
    root.mainloop()


if __name__ == "__main__":
    main()
