[![ko-fi](https://ko-fi.com/img/githubbutton_sm.svg)](https://ko-fi.com/K3K314GUP)

# Download prebuild executable
[Download v0.0.3](https://github.com/Crypto90/WindowManager/releases/download/0.0.3/Crypto90s_WindowManager.exe)

# WindowManager
A desktop app built with tkinter that lets users manage window positions and states. It uses psutil, pygetwindow, and win32gui to interact with system windows, saving and restoring positions across monitors. It also ensures that not running processes are automatically started.

![til](./preview.png)


# Requires
pip install psutil

pip install pygetwindow

pip install screeninfo

pip install pywin32

pip install win32gui

# Build executable
pyinstaller --onefile --noconsole Crypto90s_WindowManager.py
