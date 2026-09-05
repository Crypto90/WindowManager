[![ko-fi](https://ko-fi.com/img/githubbutton_sm.svg)](https://ko-fi.com/K3K314GUP)

# Crypto90's WindowManager

A fast, multithreaded desktop application built with Python and Tkinter to manage, save, and restore the exact positions, sizes, and states of application windows across multiple monitors.

WindowManager monitors your workspace layout, ensures required processes are active (launching them automatically if closed), and positions windows according to your configured presets.

---

## ⚡ What's New in v0.2.0

- **Non-Blocking Background Threading**: Window launching and ordering now run on a dedicated worker thread with live timestamped logs and a real-time "Cancel" button. The GUI remains 100% fluid with zero `(Not Responding)` freezes.
- **Human-Readable JSON Storage**: Presets are saved as clean JSON (`window_states_X.json`) with seamless automatic migration from legacy `.pkl` files.
- **Offline App Preservation**: Saving a preset no longer deletes closed applications. Offline apps remain safely stored in your preset.
- **Per-Monitor High-DPI Awareness**: Accurate coordinate positioning across mixed-DPI multi-monitor displays on Windows 10/11.
- **Maximized Window Support**: Preserves both normal restored geometry and true maximized (`SW_MAXIMIZE`) window states.
- **Custom Preset Renaming**: Rename presets (e.g. *"Work / Dev"*, *"Trading"*, *"Gaming Setup"*) directly in the UI.
- **Headless / Silent CLI Mode**: Run window ordering silently in the background without opening any GUI window (`--silent`).
- **Windows Startup Toggle**: Enable or disable launch at Windows startup with one click directly in the interface.
- **Context Menu Expansion**: Right-click to add or remove launcher overrides, or jump directly to the executable location.

---

## Command Line Usage

```bash
# Launch GUI with specific preset (1-10)
python Crypto90s_WindowManager.py --preset 2

# Run silently in background on startup (no GUI)
python Crypto90s_WindowManager.py --preset 1 --silent

# Check version
python Crypto90s_WindowManager.py --version
```

---

## Download Pre-built Executable

- 🚀 **[Download WindowManager v0.2.0 (Crypto90s_WindowManager.exe)](https://github.com/Crypto90/WindowManager/releases/download/0.2.0/Crypto90s_WindowManager.exe)**
- All releases and changelogs: [GitHub Releases](https://github.com/Crypto90/WindowManager/releases)

---

## Screenshot

![WindowManager Preview](./preview.png)

---

## Features

- **Window State Management**: Saves and restores exact coordinates, dimensions, minimized status, and maximized states.
- **Multi-Monitor Layouts**: Automatically groups and places windows on the correct monitor.
- **Process Auto-Launch**: Detects missing applications and launches them before arranging windows.
- **UWP & Windows Store App Support**: Resolves AppUserModelId (AUMID) to properly launch Windows Store and modern applications.
- **Preset Management**: Manage up to 10 independent workspace configurations with custom naming.
- **Auto-Close Timer**: Optionally closes the application automatically after ordering completes, with a cancel option.
- **Color-Coded Live Logs**: High-contrast, timestamped logs with dedicated colors for errors, warnings, info, and success.
- **File Explorer Window Support**: Manages genuine File Explorer folder windows while excluding desktop/taskbar system shells.

---

## Installation & Requirements

### Clone and Install Dependencies

```bash
git clone https://github.com/Crypto90/WindowManager.git
cd WindowManager

pip install -r requirements.txt
```

### Dependencies
- `psutil`
- `pygetwindow`
- `screeninfo`
- `pywin32` (on Windows)

---

## Build Executable with PyInstaller

```bash
pyinstaller --onefile --noconsole Crypto90s_WindowManager.py
```

---

## License

Distributed under the Apache 2.0 License. See `LICENSE` for details.
