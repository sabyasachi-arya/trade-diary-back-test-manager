Trade Diary — Backtest Manager

A fully offline Windows desktop application for logging, analysing, and managing trading backtest data. Built with Python and PyQt6.

Features
- Log trades with date, index, direction, result, and points captured
- Auto-detects the day of the week from any date
- Calendar picker for date selection
- Find, edit, and remove any trade by date
- Live sync with Excel every change instantly updates `Trade_Diary_Backtest_1.xlsx`
- Excel file saved as a proper formatted Table (TableStyleLight13)
- Analysis dashboard with:
  - Pie charts: Profit/Loss/CTC ratio, Rules vs No Rules, Trading vs Non-Trading days
  - Bar charts: Points per trade coloured by result, Weekly performance by day of week
- P/L/CTC ratio tracker
- Fully offline no internet, no server, no cloud

Tech Stack
- Python 3
- PyQt6 (native Windows GUI)
- Pandas + OpenPyXL (Excel read/write)
- Matplotlib (charts and analysis)
- PyInstaller (packaged as standalone .exe)

How to Run
1. Install Python from python.org (check "Add to PATH")
2. Double-click `Launch_Trade_Diary.bat`

Build standalone .exe
Right-click `Build_EXE.ps1` → Run with PowerShell

Screenshots
<img width="1920" height="1031" alt="Project image 4" src="https://github.com/user-attachments/assets/ba5d6b03-00a4-4680-a66f-a3417d1e11cc" />

<img width="1920" height="1029" alt="Project Image 4-2" src="https://github.com/user-attachments/assets/18b20c72-e75a-4381-af1e-6b2c49e5adcd" />

<img width="1920" height="1028" alt="Project Image 4-3" src="https://github.com/user-attachments/assets/bb947aed-bc80-4736-a69f-78969c90d372" />

<img width="1920" height="1029" alt="Project Image 4-4" src="https://github.com/user-attachments/assets/d039de70-1b6a-4534-a218-99ec267c1d50" />

