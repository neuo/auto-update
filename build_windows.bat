@echo off
setlocal

if not exist .venv (
  py -3 -m venv .venv
)

call .venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt pyinstaller

pyinstaller --noconfirm --clean --onefile --windowed --name STCloseUpdater --collect-all akshare --collect-all py_mini_racer --collect-all curl_cffi st_close_gui.py

echo.
echo Build complete. EXE: dist\STCloseUpdater.exe
pause
