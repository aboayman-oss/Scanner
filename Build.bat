@echo off
setlocal

REM Clean previous PyInstaller outputs
if exist dist rmdir /s /q dist
if exist build rmdir /s /q build

REM --noconfirm: overwrite existing artifacts without prompting
REM --onefile: bundle everything into a single executable
REM --add-data: embed the assets directory inside the bundle
REM --hidden-import: ensure openpyxl is available when packaging Excel support
pyinstaller --noconfirm --onefile ^
    --add-data "assets;assets" ^
    --hidden-import openpyxl ^
    src\main.py

if errorlevel 1 (
    echo PyInstaller build failed.
    exit /b %errorlevel%
)

echo.
echo Build complete! Check the dist folder for the executable.
pause
