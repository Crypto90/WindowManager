@echo off
echo ===================================================
echo Building Crypto90's Workspace Manager Executable
echo ===================================================

echo Installing / checking dependencies...
pip install -r requirements.txt
pip install pyinstaller

echo Building standalone Windows executable...
pyinstaller --onefile --noconsole --name "Crypto90s_WorkspaceManager" Crypto90s_WorkspaceManager.py

if %ERRORLEVEL% equ 0 (
    echo.
    echo ===================================================
    echo BUILD SUCCESSFUL!
    echo Binary: dist\Crypto90s_WorkspaceManager.exe
    echo ===================================================
) else (
    echo.
    echo BUILD FAILED! Check the error output above.
)

pause
