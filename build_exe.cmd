@echo off
setlocal

rem Read version from copilots_app/__init__.py
for /f "delims=" %%v in ('python -c "from copilots_app import __version__; print(__version__)"') do set APP_VERSION=%%v
set EXE_NAME=Copilots v%APP_VERSION%

echo Building %EXE_NAME% EXE with PyInstaller (pywebview backend)...
python -m PyInstaller --noconfirm --onefile --windowed --name "%EXE_NAME%" --icon "assets/icons/copilots.ico" --add-data "assets;assets" --add-data "copilots_app/web;copilots_app/web" --add-data "copilots_app/prompts;copilots_app/prompts" app.py
set BUILD_ERRORLEVEL=%ERRORLEVEL%

if %BUILD_ERRORLEVEL% equ 0 (
    echo.
    echo Build successful! Executable is located in "dist\%EXE_NAME%.exe"
) else (
    echo.
    echo Build failed with error code %BUILD_ERRORLEVEL%.
)

if not defined CI (
    pause
)

exit /b %BUILD_ERRORLEVEL%
