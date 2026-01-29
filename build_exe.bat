@echo off
echo Dang tao file EXE...
    python -m PyInstaller --noconfirm --onefile --windowed --name "Tool_Muasamcong_Scrape_Data" --collect-all customtkinter --collect-all playwright --collect-all pandas --collect-all openpyxl --hidden-import=requests --add-data "DATA_SAMPLE;DATA_SAMPLE" --add-data "Image;Image" --icon="Image/BST_Pharma_ICO.ico" gui_tool.py
echo.
echo Xong! File EXE nam trong thu muc 'dist'.
pause
