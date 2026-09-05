@echo off
setlocal

echo Building Copilots EXE with PyInstaller (pywebview backend)...
py -m PyInstaller --noconfirm --onefile --windowed --name "Copilots" --icon "assets/icons/copilots.png" --add-data "assets;assets" --add-data "copilots_app/web;copilots_app/web" --add-data "copilots_app/prompts;copilots_app/prompts" app.py

if %ERRORLEVEL% equ 0 (
    echo.
    echo Build successful! Executable is located in dist\Copilots.exe
) else (
    echo.
    echo Build failed with error code %ERRORLEVEL%.
)

pause
