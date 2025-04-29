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

# Class for window state
class WindowState:
    def __init__(self, process_name, position, size, process_path=None):
        self.process_name = process_name
        self.position = position
        self.size = size
        self.process_path = process_path

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
        self.log(f"Crypto's WindowManager © Crypto90. All rights reserved.")
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
                        label = f"{star}{pname} (Title: {win.title[:30]}, Size: {width}x{height}, Position: {left},{top})"
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

                state = WindowState(pname, (win.left, win.top), (win.width, win.height), process_path)
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
            running = is_process_running(state.process_path)
            if not running and state.process_path:
                try:
                    subprocess.Popen(state.process_path)
                    self.log(f"Started: {state.process_path}")
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
    root.geometry("400x400")
    app = WindowManagerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
