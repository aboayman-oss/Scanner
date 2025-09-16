@echo off
REM — Clean previous builds
if exist "Session scanner.exe" del /q "Session scanner.exe"
if exist "Session scanner.spec" del /q "Session scanner.spec"
if exist build rmdir /s /q build

REM — Run PyInstaller, output EXE directly to the base folder
pyinstaller --noconfirm --onefile --windowed ^
  --name "Session scanner" ^
  --icon "assets\thumbnail.ico" ^
  --add-data "assets;assets" ^
  --distpath . ^
  src\main.py

echo.
echo Build complete! Your EXE is: Session scanner.exe
pause