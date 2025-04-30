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


import ctypes
from ctypes import wintypes

current_version = "v0.0.4"

# Class for window state
class WindowState:
    def __init__(self, process_name, position, size, process_path=None, url=None, path=None):
        self.process_name = process_name
        self.position = position
        self.size = size
        self.process_path = process_path
        self.url = url
        self.path = path

# Load/save state
def load_window_states(filename="window_states.pkl"):
    try:
        with open(filename, "rb") as f:
            return pickle.load(f)
    except (FileNotFoundError, EOFError):
        return {}

def save_window_states(window_states, filename="window_states.pkl"):
    with open(filename, "wb") as f:
        pickle.dump(window_states, f)

def get_monitor_for_window(win, monitors):
    wx, wy = win.left + win.width // 2, win.top + win.height // 2
    for monitor in monitors:
        if monitor.x <= wx <= monitor.x + monitor.width and monitor.y <= wy <= monitor.y + monitor.height:
            return monitor
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
        if not win32gui.IsWindowVisible(hWnd) or not win32gui.IsWindowEnabled(hWnd):
            continue
        if win32gui.GetWindow(hWnd, win32con.GW_OWNER):
            continue
        if win32gui.GetWindowTextLength(hWnd) == 0:
            continue
        process_name = get_process_name_for_window(window)
        if not process_name:
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
        #print(f"Searching for identifier: {identifier}")
        
        # Run PowerShell command to get all start apps with full AppIDs using Format-List
        process = subprocess.Popen(
            ["powershell", "-Command", "Get-StartApps | Format-List Name, AppID"],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE
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
        #print("searching: " + identifier)
        for line in lines:
            if line.startswith("Name"):
                name = line.split(":", 1)[1].strip()
                #print(name)
            elif line.startswith("AppID"):
                appid = line.split(":", 1)[1].strip()
                #print("in: " + appid)
                #print("-----")
                if identifier.lower() in appid.lower():
                    return name  # or return (name, appid) if both are useful

        return None  # no match found

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
        self.window_states = load_window_states()
        self.window_mapping = []
        
        # Set minimum width and height for the main window
        self.root.minsize(width=400, height=400)
        
        
        # Process listbox with scrollbar
        self.process_listbox_frame = tk.Frame(self.root)
        self.process_listbox_frame.pack(pady=10, fill=tk.BOTH, expand=True)
        
        self.process_listbox = tk.Listbox(self.process_listbox_frame, selectmode=tk.MULTIPLE)
        self.process_listbox_scrollbar = tk.Scrollbar(self.process_listbox_frame, orient="vertical", command=self.process_listbox.yview)
        
        self.process_listbox.config(yscrollcommand=self.process_listbox_scrollbar.set)
        
        self.process_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.process_listbox_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Buttons
        tk.Button(self.root, text="Refresh Process List", command=self.refresh_window_list).pack(pady=5)
        tk.Button(self.root, text="Save", command=self.save_window_positions).pack(pady=5)
        tk.Button(self.root, text="Order", command=self.stream_order).pack(pady=5)
        tk.Button(self.root, text="Buy me a Coffee ☕", fg="white", bg="#29abe0", command=lambda: webbrowser.open("https://ko-fi.com/crypto90")).pack(pady=5)


        # Log box with scrollbar and fixed height
        self.log_text_frame = tk.Frame(self.root)
        self.log_text_frame.pack(pady=10, fill=tk.X, expand=False, side=tk.BOTTOM)  # Ensure it's at the bottom

        # Create the Text widget for the log box
        self.log_text = tk.Text(self.log_text_frame, height=10, state=tk.DISABLED, bg="black", fg="white")

        # Create the Scrollbar widget for the log box
        self.log_text_scrollbar = tk.Scrollbar(self.log_text_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.config(yscrollcommand=self.log_text_scrollbar.set)

        # Pack the Text widget to expand horizontally (fill X) but keep the fixed height
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Pack the Scrollbar widget to the right side with vertical filling (fill Y)
        self.log_text_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Adding copyright and thanks message to the log output
        self.log(f"------------------------------------------------")
        self.log(f"Crypto's WindowManager {current_version} © Crypto90. All rights reserved.")
        self.log(f"Thanks for using! Find the source code here:\nhttps://github.com/Crypto90/WindowManager")
        self.log(f"------------------------------------------------")

        self.populate_window_list()

        # If window states are loaded, restore main window's position and size
        if self.window_states:
            self.restore_main_window_position()

    def log(self, message):
        self.log_text.config(state=tk.NORMAL)
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
                label = f"★ {pname} (Not running)"
                self.process_listbox.insert(tk.END, label)
                index = self.process_listbox.size() - 1
                self.process_listbox.itemconfig(index, {'fg': 'red'})
                self.window_mapping.append(None)
        
        grouped = get_visible_windows_grouped_by_monitor()
        for monitor, items in grouped.items():
            self.process_listbox.insert(tk.END, f"=== Monitor: {monitor} ===")
            self.process_listbox.itemconfig(tk.END, {'fg': 'blue'})
            self.window_mapping.append(None)
            for win, pname in items:
                saved = pname in self.window_states
                star = "★ " if saved else "  "
                try:
                    window = gw.getWindowsWithTitle(win.title)[0]  # Get window with title
                    width = window.width
                    height = window.height
                    left = window.left
                    top = window.top

                    # Only show windows that aren't 0x0
                    if width > 0 and height > 0:
                        is_uwp = is_uwp_window(win)
                        uwp_tag = " [UWP]" if is_uwp else ""
                        label = f"{star}{pname} (Title: {win.title[:30]}, Size: {width}x{height}, Position: {left},{top}){uwp_tag}"
                    else:
                        continue  # Skip windows with size 0x0

                except Exception as e:
                    label = f"{star}{pname} (Title: {win.title[:30]}, Size: N/A, Position: N/A)"
                    print(f"Error retrieving window size and position: {e}")

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
        new_states = {}

        for idx in selected_indices:
            item = self.window_mapping[idx]
            if item:
                win, pname = item
                try:
                    _, pid = win32process.GetWindowThreadProcessId(win._hWnd)
                    process = psutil.Process(pid)
                    process_path = process.exe()
                except Exception:
                    process_path = None

                # Capture additional information based on the process name
                url = None
                path = None
                if "chrome.exe" in pname.lower():
                    # Try to capture the URL from Chrome
                    # This assumes the browser has a specific way to get the URL via its window
                    url = self.get_chrome_url(pname)  # You will need to define this method to retrieve the URL
                elif "explorer.exe" in pname.lower():
                    # Capture the current directory for Explorer
                    path = self.get_explorer_path(pname)  # Define this method to capture Explorer's current path

                state = WindowState(pname, (win.left, win.top), (win.width, win.height), process_path, url, path)
                new_states[pname] = state

        # Always save the main window's position
        main_window_position = self.root.winfo_x(), self.root.winfo_y()
        main_window_size = self.root.winfo_width(), self.root.winfo_height()
        new_states["main_window"] = WindowState("main_window", main_window_position, main_window_size)

        # Filter out non-running processes except "main_window"
        cleaned_states = {}
        for pname, state in new_states.items():
            if pname == "main_window" or is_process_running(state.process_path):
                cleaned_states[pname] = state
            else:
                self.log(f"Removed non-running process from saved state: {pname}")

        self.window_states = cleaned_states
        save_window_states(self.window_states)
        self.window_states = load_window_states()
        self.refresh_window_list()
        self.log("Window positions saved.")

    def get_chrome_url(self, pname):
        """Method to extract the URL from Chrome window."""
        # Implement logic to retrieve the URL of the active Chrome tab (This may involve inspecting Chrome's internal state or using a Chrome extension API)
        # For simplicity, let's return a dummy URL
        return "https://www.example.com"  # Replace with actual method to extract Chrome URL


    def get_explorer_path(self, pname):
        """Method to extract the current path from Explorer."""
        # For Windows Explorer, we could try to use the shell to get the active directory or window path.
        # Below is just a placeholder example.
        if pname.lower() == "explorer.exe":
            return os.getcwd()  # Return the current working directory for simplicity
        return None
    
    def restore_main_window_position(self):
        if "main_window" in self.window_states:
            state = self.window_states["main_window"]
            self.root.geometry(f"{state.size[0]}x{state.size[1]}+{state.position[0]}+{state.position[1]}")
            self.log("Restored main window position and size.")


    def stream_order(self):
        started_any = False

        for pname, state in self.window_states.items():
            if not hasattr(state, "process_path") or state.process_path is None:
                continue  # Skip entries like "main_window" or invalid ones
            #print(state.process_path)
            running = is_process_running(state.process_path)
            if not running and state.process_path:
                try:
                    # Check if the app is a UWP app by looking for "WindowsApps" in the path
                    if "WindowsApps" in state.process_path:
                        # Handle UWP app start using AppUserModelId or AppExecutionAlias
                        self.log(f"Starting UWP app: {pname}")

                        # Extract the app name from the path and remove the ".exe" extension if present
                        app_name = state.process_path.split("\\")[-2]  # Extract the last part (the app name)
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
                            #self.log(f"Started UWP app: {uwp_app_name}")
                        else:
                            self.log(f"Failed to find AppUserModelId for {pname}")

                    else:
                        # For Win32 apps, use Popen
                        subprocess.Popen(state.process_path)
                        self.log(f"Started Win32 app: {state.process_path}")

                    started_any = True
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to start '{pname}': {e}")
                    self.log(f"Error: Failed to start '{pname}': {e}")

        if started_any:
            self.root.update()
            self.log("Waiting 5 seconds for windows to appear...")
            time.sleep(5)
            # Refresh the process list after the wait
            self.refresh_window_list()

        for pname, state in self.window_states.items():
            if pname == "main_window":
                continue  # Skip placeholder entries
            
            windows = [w for w in gw.getAllWindows()
                       if get_process_name_for_window(w) == pname and win32gui.IsWindowVisible(w._hWnd)]
            
            if windows:
                win = windows[0]
                win.moveTo(*state.position)
                win.resizeTo(*state.size)
                self.log(f"Moved and resized window for {pname}")
            else:
                self.log(f"Warning: Window for '{pname}' not found.")
                # Start the process if the window is not found
                if state.process_path:
                    try:
                        subprocess.Popen(state.process_path)
                        self.log(f"Started: {state.process_path}")
                    except Exception as e:
                        self.log(f"Error: Failed to start '{pname}': {e}")
                messagebox.showwarning("Warning", f"Window for '{pname}' not found. Starting process.")



def main():
    root = tk.Tk()
    root.geometry("600x500")
    app = WindowManagerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
