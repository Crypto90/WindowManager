import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
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
import shutil

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

current_version = "v0.2.2"


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
    """
    Returns a reliable, writable directory for presets and config.
    Automatically migrates data from legacy WindowManager directories if present.
    """
    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    # Test if the directory is writable (portable mode)
    test_path = os.path.join(base_dir, ".wm_write_test")
    try:
        with open(test_path, "w") as f:
            f.write("ok")
        os.remove(test_path)
        return base_dir
    except (PermissionError, OSError):
        # Fallback to %LOCALAPPDATA%\Crypto90s_WorkspaceManager
        appdata = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
        target_dir = os.path.join(appdata, "Crypto90s_WorkspaceManager")
        legacy_dir = os.path.join(appdata, "Crypto90s_WindowManager")

        os.makedirs(target_dir, exist_ok=True)

        # Migrate presets from legacy directory if needed
        if os.path.isdir(legacy_dir) and not os.listdir(target_dir):
            try:
                for item in os.listdir(legacy_dir):
                    s = os.path.join(legacy_dir, item)
                    d = os.path.join(target_dir, item)
                    if os.path.isfile(s):
                        shutil.copy2(s, d)
                print("Info: Migrated preset files from legacy WindowManager directory.")
            except Exception as e:
                print(f"Warning: Could not copy legacy presets: {e}")

        return target_dir


def parse_args():
    parser = argparse.ArgumentParser(description="Crypto90's Workspace Manager Preset Selector")
    parser.add_argument("--preset", type=int, default=1, choices=range(1, 11),
                        help="Preset number to load (1-10, default: 1)")
    parser.add_argument("--silent", "--headless", action="store_true", dest="silent",
                        help="Run ordering in background without opening the GUI")
    parser.add_argument("--version", action="version", version=f"Crypto90's Workspace Manager {current_version}")
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


def get_preset_filepath(preset_number, ext=".json", prefix="workspace_states_"):
    data_dir = get_data_dir()
    return os.path.join(data_dir, f"{prefix}{preset_number}{ext}")


def load_window_states(preset_number=1):
    """
    Loads window states for a preset.
    Checks workspace_states_{n}.json first, then falls back to legacy files.
    """
    candidates = [
        get_preset_filepath(preset_number, ".json", "workspace_states_"),
        get_preset_filepath(preset_number, ".json", "window_states_"),
        get_preset_filepath(preset_number, ".pkl", "workspace_states_"),
        get_preset_filepath(preset_number, ".pkl", "window_states_"),
    ]

    for path in candidates:
        if not os.path.isfile(path):
            continue

        if path.endswith(".json"):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict):
                    raw_states = data.get("workspace_states", {}) or data.get("window_states", {})
                    config = data.get("config", {})
                    window_states = {}
                    for k, v in raw_states.items():
                        state_obj = WindowState.from_dict(v)
                        if state_obj:
                            window_states[k] = state_obj
                    return window_states, config
            except Exception as e:
                print(f"Warning: Failed to load JSON preset from {path}: {e}")

        elif path.endswith(".pkl"):
            try:
                with open(path, "rb") as f:
                    data = pickle.load(f)
                if isinstance(data, dict):
                    raw_states = data.get("workspace_states", {}) or data.get("window_states", {})
                    config = data.get("config", {})
                    if not raw_states and not config:
                        raw_states = data
                    
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

                    save_window_states(window_states, config, preset_number)
                    print(f"Info: Migrated preset {preset_number} from legacy PKL to JSON format.")
                    return window_states, config
            except Exception as e:
                print(f"Warning: Failed to load PKL preset from {path}: {e}")

    return {}, {}


def save_window_states(window_states, config=None, preset_number=1):
    """Saves window states and configuration to a clean JSON file."""
    json_path = get_preset_filepath(preset_number, ".json", "workspace_states_")
    serialized_states = {}
    for k, v in window_states.items():
        if isinstance(v, WindowState):
            serialized_states[k] = v.to_dict()
        elif isinstance(v, dict):
            serialized_states[k] = v

    data = {
        "version": current_version,
        "workspace_states": serialized_states,
        "window_states": serialized_states,  # Backward compatibility
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
    Filters out shell/desktop windows and Workspace Manager itself, while preserving File Explorer.
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

        # Exclude Workspace Manager itself
        if pid == current_pid:
            continue

        # For explorer.exe, allow CabinetWClass (actual folder windows), ignore Desktop/Taskbar
        if process_name.lower() == "explorer.exe":
            class_name = win32gui.GetClassName(hWnd)
            if class_name != "CabinetWClass":
                continue

        if not win32gui.IsWindowEnabled(hWnd):
            continue

        if not (win32gui.IsWindowVisible(hWnd) or win32gui.IsIconic(hWnd)):
            continue

        if win32gui.GetWindow(hWnd, win32con.GW_OWNER):
            continue

        if win32gui.GetWindowTextLength(hWnd) == 0:
            continue

        monitor = get_monitor_for_window(window, monitors)
        if monitor:
            mon_id = f"Monitor {monitor.x}x{monitor.y} ({monitor.width}x{monitor.height})"
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
    """Extracts package folder name and looks up AUMID for UWP applications."""
    if not path or "WindowsApps" not in path:
        return None, None
    try:
        parts = path.split("\\")
        for part in parts:
            if "__" in part:
                package_folder = part
                identifier = package_folder.split("__")[-1]
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
    """Checks if Workspace Manager is configured to run at Windows startup."""
    if not IS_WINDOWS:
        return False
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_READ)
        val = None
        for k in ("Crypto90s_WorkspaceManager", "Crypto90s_WindowManager"):
            try:
                val, _ = winreg.QueryValueEx(key, k)
                if val:
                    break
            except FileNotFoundError:
                continue
        winreg.CloseKey(key)
        return bool(val)
    except Exception:
        return False


def set_start_with_windows(enabled, preset=1):
    """Enables or disables Workspace Manager launch at Windows startup."""
    if not IS_WINDOWS:
        return False
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_SET_VALUE)
        if enabled:
            if getattr(sys, "frozen", False):
                exe_cmd = f'"{sys.executable}" --preset {preset} --silent'
            else:
                exe_cmd = f'"{sys.executable}" "{os.path.abspath(__file__)}" --preset {preset} --silent'
            winreg.SetValueEx(key, "Crypto90s_WorkspaceManager", 0, winreg.REG_SZ, exe_cmd)
            try:
                winreg.DeleteValue(key, "Crypto90s_WindowManager")
            except FileNotFoundError:
                pass
        else:
            for k in ("Crypto90s_WorkspaceManager", "Crypto90s_WindowManager"):
                try:
                    winreg.DeleteValue(key, k)
                except FileNotFoundError:
                    pass
        winreg.CloseKey(key)
        return True
    except Exception as e:
        print(f"Error setting Windows startup registry: {e}")
        return False


class WorkspaceManagerApp:
    def __init__(self, root, initial_preset=1):
        self.root = root
        self.root.title("Crypto90's Workspace Manager")
        self.current_preset_num = initial_preset

        self.window_states, self.config = load_window_states(self.current_preset_num)
        self.tree_item_map = {}  # item_id -> (state_key, window_obj, pname, ppath, is_offline)
        self.ordering_in_progress = False
        self.cancel_ordering = False
        self.auto_close_job = None

        # Modern Dark Slate Theme Palette
        self.BG_MAIN = "#121418"
        self.BG_CARD = "#1a1d24"
        self.BG_HOVER = "#242832"
        self.BORDER_COLOR = "#2a303c"
        self.ACCENT_CYAN = "#00d2ff"
        self.ACCENT_GREEN = "#10b981"
        self.ACCENT_RED = "#ef4444"
        self.TEXT_PRIMARY = "#f1f5f9"
        self.TEXT_MUTED = "#94a3b8"

        self.root.minsize(width=780, height=560)
        self.root.geometry("860x640")
        self.root.configure(bg=self.BG_MAIN)

        self._configure_ttk_styles()
        self._build_header_ui()
        self._build_preset_bar_ui()
        self._build_options_ui()
        self._build_table_ui()
        self._build_actions_ui()
        self._build_console_ui()
        self._build_status_bar_ui()
        self._build_context_menu()

        self.update_preset_labels()
        self.populate_window_list()

        if "main_window" in self.window_states:
            self.restore_main_window_position()

        # Initial launch trigger
        self.root.after(350, self.start_stream_order)

    def _configure_ttk_styles(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        font_family = "Segoe UI" if IS_WINDOWS else "Helvetica"

        # Treeview Dark Styling
        style.configure(
            "Custom.Treeview",
            background="#16191f",
            foreground=self.TEXT_PRIMARY,
            fieldbackground="#16191f",
            rowheight=26,
            font=(font_family, 9),
            borderwidth=0
        )
        style.configure(
            "Custom.Treeview.Heading",
            background="#20242d",
            foreground=self.TEXT_MUTED,
            font=(font_family, 9, "bold"),
            relief="flat",
            padding=5
        )
        style.map(
            "Custom.Treeview",
            background=[("selected", "#0d9488")],  # Sleek teal highlight
            foreground=[("selected", "#ffffff")]
        )
        style.map(
            "Custom.Treeview.Heading",
            background=[("active", "#282e39")]
        )

    def _build_header_ui(self):
        header_frame = tk.Frame(self.root, bg=self.BG_CARD, height=44)
        header_frame.pack(fill=tk.X, padx=0, pady=0)

        # Brand / Title
        brand_label = tk.Label(
            header_frame,
            text="🖥️  Crypto90's Workspace Manager",
            font=("Segoe UI" if IS_WINDOWS else "Helvetica", 11, "bold"),
            fg=self.TEXT_PRIMARY,
            bg=self.BG_CARD
        )
        brand_label.pack(side=tk.LEFT, padx=14, pady=8)

        # Version Badge
        version_badge = tk.Label(
            header_frame,
            text=current_version,
            font=("Segoe UI" if IS_WINDOWS else "Helvetica", 8, "bold"),
            fg=self.ACCENT_CYAN,
            bg="#222834",
            padx=6,
            pady=1
        )
        version_badge.pack(side=tk.LEFT, padx=4, pady=8)

        # Display Counter
        monitors_count = len(get_monitors()) if IS_WINDOWS else 1
        self.display_badge = tk.Label(
            header_frame,
            text=f"🖥️ Displays Detected: {monitors_count}",
            font=("Segoe UI" if IS_WINDOWS else "Helvetica", 9),
            fg=self.TEXT_MUTED,
            bg=self.BG_CARD
        )
        self.display_badge.pack(side=tk.RIGHT, padx=14, pady=8)

    def _build_preset_bar_ui(self):
        preset_frame = tk.Frame(self.root, bg=self.BG_MAIN)
        preset_frame.pack(fill=tk.X, padx=14, pady=(10, 4))

        tk.Label(
            preset_frame,
            text="Active Preset:",
            fg=self.TEXT_PRIMARY,
            bg=self.BG_MAIN,
            font=("Segoe UI" if IS_WINDOWS else "Helvetica", 9, "bold")
        ).pack(side=tk.LEFT, padx=(0, 8))

        self.current_preset = tk.StringVar(value=f"Preset {self.current_preset_num}")
        self.preset_radios = []

        # Segmented Preset Buttons
        p_container = tk.Frame(preset_frame, bg=self.BG_CARD, bd=1, relief=tk.FLAT)
        p_container.pack(side=tk.LEFT, padx=0)

        for i in range(1, 11):
            name = f"Preset {i}"
            rb = tk.Radiobutton(
                p_container,
                text=str(i),
                variable=self.current_preset,
                value=name,
                command=self.switch_preset,
                bg=self.BG_CARD,
                fg=self.TEXT_MUTED,
                selectcolor="#0f766e",
                activebackground=self.BG_CARD,
                activeforeground=self.ACCENT_CYAN,
                font=("Segoe UI" if IS_WINDOWS else "Helvetica", 9),
                indicatoron=False,
                padx=6,
                pady=3,
                bd=0
            )
            rb.pack(side=tk.LEFT, padx=1)
            self.preset_radios.append(rb)

        # Rename Button
        tk.Button(
            preset_frame,
            text="✏️ Rename",
            bg="#2a303c",
            fg=self.TEXT_PRIMARY,
            activebackground=self.BG_HOVER,
            activeforeground="#ffffff",
            relief=tk.FLAT,
            padx=8,
            pady=3,
            font=("Segoe UI" if IS_WINDOWS else "Helvetica", 8, "bold"),
            command=self.rename_current_preset
        ).pack(side=tk.LEFT, padx=10)

    def _build_options_ui(self):
        options_frame = tk.Frame(self.root, bg=self.BG_MAIN)
        options_frame.pack(fill=tk.X, padx=14, pady=(2, 6))

        self.auto_close_var = tk.BooleanVar(value=self.config.get("auto_close", False))
        self.auto_close_checkbox = tk.Checkbutton(
            options_frame,
            text="Auto-close after ordering",
            variable=self.auto_close_var,
            command=self.on_auto_close_toggle,
            bg=self.BG_MAIN,
            fg=self.TEXT_MUTED,
            selectcolor="#2a303c",
            activebackground=self.BG_MAIN,
            activeforeground=self.TEXT_PRIMARY,
            font=("Segoe UI" if IS_WINDOWS else "Helvetica", 9)
        )
        self.auto_close_checkbox.pack(side=tk.LEFT, padx=(0, 16))

        self.startup_var = tk.BooleanVar(value=is_start_with_windows_enabled())
        self.startup_checkbox = tk.Checkbutton(
            options_frame,
            text="Run at Windows Startup",
            variable=self.startup_var,
            command=self.on_startup_toggle,
            bg=self.BG_MAIN,
            fg=self.TEXT_MUTED,
            selectcolor="#2a303c",
            activebackground=self.BG_MAIN,
            activeforeground=self.TEXT_PRIMARY,
            font=("Segoe UI" if IS_WINDOWS else "Helvetica", 9)
        )
        self.startup_checkbox.pack(side=tk.LEFT, padx=0)

    def _build_table_ui(self):
        table_card = tk.Frame(self.root, bg=self.BG_CARD, bd=1, relief=tk.FLAT)
        table_card.pack(fill=tk.BOTH, expand=True, padx=14, pady=4)

        # Columns definition
        columns = ("status", "app", "title", "monitor", "coords", "state")
        self.tree = ttk.Treeview(
            table_card,
            columns=columns,
            show="headings",
            selectmode="extended",
            style="Custom.Treeview"
        )

        self.tree.heading("status", text="Status", anchor=tk.W)
        self.tree.heading("app", text="Application", anchor=tk.W)
        self.tree.heading("title", text="Window Title", anchor=tk.W)
        self.tree.heading("monitor", text="Monitor", anchor=tk.W)
        self.tree.heading("coords", text="Coordinates / Size", anchor=tk.W)
        self.tree.heading("state", text="Window State", anchor=tk.CENTER)

        self.tree.column("status", width=95, minwidth=70, stretch=False)
        self.tree.column("app", width=140, minwidth=110, stretch=False)
        self.tree.column("title", width=220, minwidth=140, stretch=True)
        self.tree.column("monitor", width=150, minwidth=100, stretch=False)
        self.tree.column("coords", width=140, minwidth=110, stretch=False)
        self.tree.column("state", width=95, minwidth=70, stretch=False)

        # Scrollbar
        sb_y = tk.Scrollbar(table_card, orient=tk.VERTICAL, command=self.tree.yview, bg=self.BG_CARD)
        self.tree.configure(yscrollcommand=sb_y.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb_y.pack(side=tk.RIGHT, fill=tk.Y)

        # Row Tags
        self.tree.tag_configure("saved", foreground="#34d399")       # Mint / Emerald
        self.tree.tag_configure("override", foreground="#38bdf8")    # Cyan
        self.tree.tag_configure("offline", foreground="#f87171")     # Light Red
        self.tree.tag_configure("normal", foreground="#e2e8f0")      # Bright White
        self.tree.tag_configure("header", background="#222834", foreground="#38bdf8", font=("Segoe UI" if IS_WINDOWS else "Helvetica", 9, "bold"))

    def _build_actions_ui(self):
        action_bar = tk.Frame(self.root, bg=self.BG_MAIN)
        action_bar.pack(fill=tk.X, padx=14, pady=8)

        # Buttons with modern flat styling
        tk.Button(
            action_bar,
            text="🔄 Refresh List",
            bg="#2563eb",
            fg="white",
            activebackground="#1d4ed8",
            activeforeground="white",
            relief=tk.FLAT,
            padx=12,
            pady=4,
            font=("Segoe UI" if IS_WINDOWS else "Helvetica", 9, "bold"),
            command=self.refresh_window_list
        ).pack(side=tk.LEFT, padx=(0, 6))

        tk.Button(
            action_bar,
            text="💾 Save Selected to Preset",
            bg="#475569",
            fg="white",
            activebackground="#334155",
            activeforeground="white",
            relief=tk.FLAT,
            padx=12,
            pady=4,
            font=("Segoe UI" if IS_WINDOWS else "Helvetica", 9, "bold"),
            command=self.save_window_positions
        ).pack(side=tk.LEFT, padx=6)

        self.order_button = tk.Button(
            action_bar,
            text="🚀 Start, Resize & Order",
            bg="#059669",
            fg="white",
            activebackground="#047857",
            activeforeground="white",
            relief=tk.FLAT,
            padx=14,
            pady=4,
            font=("Segoe UI" if IS_WINDOWS else "Helvetica", 9, "bold"),
            command=self.toggle_stream_order
        )
        self.order_button.pack(side=tk.LEFT, padx=6)

        tk.Button(
            action_bar,
            text="☕ Buy Coffee",
            bg="#d97706",
            fg="white",
            activebackground="#b45309",
            activeforeground="white",
            relief=tk.FLAT,
            padx=10,
            pady=4,
            font=("Segoe UI" if IS_WINDOWS else "Helvetica", 9, "bold"),
            command=lambda: webbrowser.open("https://ko-fi.com/crypto90")
        ).pack(side=tk.RIGHT, padx=0)

    def _build_console_ui(self):
        console_container = tk.Frame(self.root, bg=self.BG_MAIN)
        console_container.pack(fill=tk.X, padx=14, pady=(2, 6))

        # Console Header
        bar = tk.Frame(console_container, bg=self.BG_CARD, height=24)
        bar.pack(fill=tk.X)

        tk.Label(
            bar,
            text="ACTIVITY & DIAGNOSTICS CONSOLE",
            font=("Segoe UI" if IS_WINDOWS else "Helvetica", 8, "bold"),
            fg=self.TEXT_MUTED,
            bg=self.BG_CARD
        ).pack(side=tk.LEFT, padx=8, pady=2)

        tk.Button(
            bar,
            text="Clear Console",
            bg=self.BG_CARD,
            fg=self.TEXT_MUTED,
            activebackground=self.BG_CARD,
            activeforeground=self.TEXT_PRIMARY,
            font=("Segoe UI" if IS_WINDOWS else "Helvetica", 7),
            relief=tk.FLAT,
            bd=0,
            command=self.clear_console
        ).pack(side=tk.RIGHT, padx=6)

        # Text Console
        log_frame = tk.Frame(console_container, bg="#0f1115")
        log_frame.pack(fill=tk.X)

        self.log_text = tk.Text(
            log_frame,
            height=6,
            state=tk.DISABLED,
            bg="#0f1115",
            fg="#e2e8f0",
            insertbackground="white",
            highlightbackground=self.BORDER_COLOR,
            font=("Consolas" if IS_WINDOWS else "Courier", 9),
            relief=tk.FLAT,
            padx=6,
            pady=4
        )

        sb = tk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview, bg="#0f1115")
        self.log_text.config(yscrollcommand=sb.set)

        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        # Tag configurations
        self.log_text.tag_config("error", foreground="#f87171")    # Red
        self.log_text.tag_config("info", foreground="#38bdf8")     # Cyan
        self.log_text.tag_config("success", foreground="#34d399")  # Emerald
        self.log_text.tag_config("warn", foreground="#fbbf24")     # Amber
        self.log_text.tag_config("time", foreground="#64748b")     # Slate timestamp

    def _build_status_bar_ui(self):
        self.status_bar = tk.Label(
            self.root,
            text="Ready",
            bg=self.BG_CARD,
            fg=self.TEXT_MUTED,
            font=("Segoe UI" if IS_WINDOWS else "Helvetica", 8),
            anchor=tk.W,
            padx=12,
            pady=3
        )
        self.status_bar.pack(fill=tk.X, side=tk.BOTTOM)

    def _build_context_menu(self):
        self.context_menu = tk.Menu(self.root, tearoff=0, bg="#20242d", fg="white", activebackground="#0d9488")
        self.context_menu.add_command(label="Add Launcher Override...", command=self.set_launcher_override)
        self.context_menu.add_command(label="Remove Launcher Override", command=self.remove_launcher_override)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Open Executable Location", command=self.open_file_location)

        self.tree.bind("<Button-3>", self.show_context_menu)
        if sys.platform == "darwin":
            self.tree.bind("<Button-2>", self.show_context_menu)

    def clear_console(self):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.config(state=tk.DISABLED)

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
                elif any(w in lower for keyword in ["success", "saved", "restored", "started", "arranged"] for w in [keyword]):
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

            # Update status bar
            self.status_bar.config(text=f"Last event: {message}")

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
            item_id = self.tree.identify_row(event.y)
            if not item_id or item_id not in self.tree_item_map:
                return
            entry = self.tree_item_map[item_id]
            if not entry:
                return

            self.tree.selection_set(item_id)
            self.context_menu.post(event.x_root, event.y_root)
        except Exception as e:
            print(f"Error showing context menu: {e}")

    def get_selected_mapping_entry(self):
        selection = self.tree.selection()
        if not selection:
            return None
        item_id = selection[0]
        return self.tree_item_map.get(item_id)

    def set_launcher_override(self):
        entry = self.get_selected_mapping_entry()
        if not entry:
            return
        state_key, _, pname, _, _ = entry

        file_path = filedialog.askopenfilename(
            title=f"Select Launcher Executable for {pname}",
            filetypes=[("Executable Files", "*.exe"), ("All Files", "*.*")]
        )
        if not file_path:
            return

        if state_key in self.window_states:
            self.window_states[state_key].launcher_override = file_path
        else:
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
        state_key, _, pname, _, _ = entry

        if state_key in self.window_states and self.window_states[state_key].launcher_override:
            self.window_states[state_key].launcher_override = None
            save_window_states(self.window_states, self.config, self.current_preset_num)
            self.log(f"Removed launcher override for {pname}", tag="info")
            self.refresh_window_list()

    def open_file_location(self):
        entry = self.get_selected_mapping_entry()
        if not entry:
            return
        state_key, _, pname, ppath, _ = entry

        path = None
        if state_key in self.window_states:
            st = self.window_states[state_key]
            path = st.launcher_override or st.process_path
        if not path:
            path = ppath

        if path and os.path.exists(path):
            if IS_WINDOWS:
                subprocess.Popen(f'explorer.exe /select,"{path}"')
            else:
                subprocess.Popen(["open", os.path.dirname(path)])
        else:
            self.log(f"Executable path not found on disk: {path}", tag="warn")

    def refresh_window_list(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.tree_item_map.clear()
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

        saved_item_ids_to_select = []

        # 1. Render Saved Applications that are NOT currently running
        for key, state in self.window_states.items():
            if key == "main_window":
                continue
            if state.process_name not in current_running_keys and key not in current_running_keys:
                status_label = "★★ Override" if state.launcher_override else "★ Offline"
                title_disp = state.window_title or "(Closed)"
                coords_disp = f"{state.position[0]}, {state.position[1]} [{state.size[0]}x{state.size[1]}]"
                state_disp = "Maximized" if state.maximized else ("Minimized" if state.minimized else "Normal")

                row_id = self.tree.insert(
                    "",
                    "end",
                    values=(status_label, state.process_name, title_disp, "Saved State", coords_disp, state_disp),
                    tags=("offline",)
                )
                self.tree_item_map[row_id] = (key, None, state.process_name, state.process_path, True)
                saved_item_ids_to_select.append(row_id)

        # 2. Render Open Windows grouped by Monitor
        for monitor_name, items in grouped.items():
            hdr_id = self.tree.insert(
                "",
                "end",
                values=(f"=== {monitor_name} ===", "", "", "", "", ""),
                tags=("header",)
            )
            self.tree_item_map[hdr_id] = None

            for win, pname, ppath in items:
                title = win.title if win else ""
                key = f"{pname}::{title[:40]}" if title else pname

                saved = key in self.window_states or pname in self.window_states
                active_state = self.window_states.get(key) or self.window_states.get(pname)
                saved_override = bool(active_state and active_state.launcher_override)

                status_label = "★★ Saved" if saved_override else ("★ Saved" if saved else "🟢 Active")
                row_tag = "override" if saved_override else ("saved" if saved else "normal")

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

                    is_uwp = is_uwp_window(win)
                    state_tokens = []
                    if hwnd_maximized:
                        state_tokens.append("Maximized")
                    elif hwnd_minimized:
                        state_tokens.append("Minimized")
                    else:
                        state_tokens.append("Normal")
                    if is_uwp:
                        state_tokens.append("UWP")

                    state_disp = " • ".join(state_tokens)
                    coords_disp = f"{left}, {top} [{width}x{height}]"
                except Exception:
                    coords_disp = "N/A"
                    state_disp = "Normal"

                mon_short = monitor_name.split()[0] + " " + monitor_name.split()[1] if " " in monitor_name else monitor_name
                row_id = self.tree.insert(
                    "",
                    "end",
                    values=(status_label, pname, title, mon_short, coords_disp, state_disp),
                    tags=(row_tag,)
                )
                self.tree_item_map[row_id] = (key, win, pname, ppath, False)

                if saved:
                    saved_item_ids_to_select.append(row_id)

        # Pre-select all saved items
        if saved_item_ids_to_select:
            self.tree.selection_set(saved_item_ids_to_select)

    def save_window_positions(self):
        """Saves selected window positions while preserving offline apps."""
        selection = self.tree.selection()
        if not selection:
            self.log("No windows selected to save.", tag="warn")
            return

        previous_states, _ = load_window_states(self.current_preset_num)
        new_window_states = {}

        for item_id in selection:
            entry = self.tree_item_map.get(item_id)
            if not entry:
                continue

            state_key, win, pname, ppath, is_offline = entry

            if is_offline or win is None:
                if state_key in previous_states:
                    new_window_states[state_key] = previous_states[state_key]
                    self.log(f"Preserved offline app in preset: {pname}", tag="info")
                continue

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
        self.order_button.config(text="⏹ Stop / Cancel", bg="#dc2626")

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

            # Phase 1: Launch missing apps
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
                self.log("Waiting for launched applications to open...", tag="info")
                time.sleep(2.0)

            # Phase 2: Poll and reposition
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

                    matched_win = None
                    for w in gw.getAllWindows():
                        pname, ppath, _ = get_process_info_for_window(w)
                        if pname == state.process_name:
                            if state.window_title and state.window_title in w.title:
                                matched_win = w
                                break
                            elif not matched_win:
                                matched_win = w

                    if matched_win:
                        hwnd = matched_win._hWnd
                        try:
                            placement = win32gui.GetWindowPlacement(hwnd)
                            if placement[1] in (win32con.SW_SHOWMINIMIZED, win32con.SW_SHOWMAXIMIZED):
                                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

                            matched_win.moveTo(*state.position)
                            matched_win.resizeTo(*state.size)

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
        self.order_button.config(text="🚀 Start, Resize & Order", bg="#059669")
        self.refresh_window_list()

        if self.auto_close_var.get() and not self.cancel_ordering:
            self.log("Auto-close is enabled. Exiting in 5 seconds... (Click 'Stop' to cancel)", tag="info")
            self.order_button.config(text="Cancel Auto-Close", bg="#d97706", command=self.cancel_auto_close)
            self.auto_close_job = self.root.after(5000, self.root.destroy)

    def cancel_auto_close(self):
        if self.auto_close_job:
            self.root.after_cancel(self.auto_close_job)
            self.auto_close_job = None
            self.log("Auto-close cancelled.", tag="info")
            self.order_button.config(text="🚀 Start, Resize & Order", bg="#059669", command=self.toggle_stream_order)

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
            if hasattr(e, 'winerror') and e.winerror == 740:
                self.log(f"Elevation required for {os.path.basename(path)}. Launching via ShellExecute...", tag="warn")
                if IS_WINDOWS:
                    ctypes.windll.shell32.ShellExecuteW(None, "open", path, None, None, 1)
            else:
                raise

    def try_start_application(self, state):
        """Starts an application, resolving UWP or relocated executables."""
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

        # 3. Fallback: Search for relocated executable
        base_dir = os.path.dirname(launcher_path)
        filename = os.path.basename(launcher_path)
        parent_dir = os.path.dirname(base_dir)

        if parent_dir and len(parent_dir.strip(":/\\")) > 1 and os.path.isdir(parent_dir):
            for root, dirs, files in os.walk(parent_dir):
                rel = os.path.relpath(root, parent_dir)
                if rel.count(os.sep) > 2:
                    continue
                if filename in files:
                    new_path = os.path.join(root, filename)
                    if os.path.isfile(new_path):
                        self.log(f"Detected updated path for {pname}: {new_path}", tag="info")
                        state.process_path = new_path
                        save_window_states(self.window_states, self.config, self.current_preset_num)
                        self.launch_independent(new_path)
                        return

        self.log(f"Could not find or launch: {launcher_path}", tag="error")


def run_headless(preset_number=1):
    """Headless CLI mode: orders windows silently in background without GUI."""
    print(f"Crypto90's Workspace Manager {current_version} [Headless Mode]")
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
    print("Headless workspace ordering complete.")


def main():
    enable_high_dpi()
    args = parse_args()

    if args.silent:
        run_headless(args.preset)
        return

    root = tk.Tk()
    app = WorkspaceManagerApp(root, initial_preset=args.preset)
    root.mainloop()


if __name__ == "__main__":
    main()
