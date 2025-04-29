# WindowManager
A desktop app built with tkinter that lets users manage window positions and states. It uses psutil, pygetwindow, and win32gui to interact with system windows, saving and restoring positions across monitors. It also ensures that not running processes are automatically started.

##Requires
pip install psutil
pip install pygetwindow
pip install screeninfo
pip install pywin32
pip install win32gui

# Build executable
pyinstaller --onefile --noconsole Crypto90s_WindowManager.py
