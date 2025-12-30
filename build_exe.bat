@echo off
echo Dang tao file EXE...
python -m PyInstaller --noconfirm --onefile --windowed --name "Tool_Scrape_Muasamcong" --collect-all customtkinter --collect-all playwright --collect-all pandas --collect-all openpyxl --icon=NONE gui_tool.py
echo.
echo Xong! File EXE nam trong thu muc 'dist'.
pause
