@echo off
REM ============================================================================
REM  build_exe.bat - package Voice-to-Word as a standalone Windows app.
REM
REM  Uses PyInstaller to bundle Python, the app and all its libraries into
REM  dist\VoiceToWord\. That folder can be copied to any Windows PC and run
REM  without installing Python. Only this build step needs Python.
REM ============================================================================

REM Work from the folder this script is in.
cd /d "%~dp0"

REM A separate environment just for building, so the .exe contains only what it needs.
python -m venv .build-venv || (echo Python 3.9+ is required: https://www.python.org/downloads/ & pause & exit /b 1)
.build-venv\Scripts\python -m pip install --upgrade pip
.build-venv\Scripts\python -m pip install -r requirements.txt pyinstaller || (pause & exit /b 1)

REM  --windowed     : no console window behind the app
REM  --onedir       : a folder of files (starts faster than a single huge .exe)
REM  --collect-all  : include data files and native libraries these packages
REM                   load at runtime, which PyInstaller cannot detect on its own
.build-venv\Scripts\pyinstaller --noconfirm --clean --windowed --onedir ^
  --name VoiceToWord ^
  --collect-all faster_whisper ^
  --collect-all ctranslate2 ^
  --collect-all av ^
  --collect-all tokenizers ^
  --collect-all onnxruntime ^
  --collect-all docx ^
  --collect-all pyaudiowpatch ^
  voice_to_word.py || (pause & exit /b 1)

echo.
echo Done! Your app is in:  dist\VoiceToWord\VoiceToWord.exe
echo Copy the whole dist\VoiceToWord folder anywhere you like.
pause
