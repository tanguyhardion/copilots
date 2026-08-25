@echo off
setlocal

echo Building Copilots EXE with PyInstaller...
py -m PyInstaller --noconfirm --onefile --windowed --name "Copilots" --icon "assets/icons/copilots.png" --add-data "assets;assets" app.py

if %ERRORLEVEL% equ 0 (
    echo.
    echo Build successful! Executable is located in dist\Copilots\Copilots.exe
) else (
    echo.
    echo Build failed with error code %ERRORLEVEL%.
)

pause
