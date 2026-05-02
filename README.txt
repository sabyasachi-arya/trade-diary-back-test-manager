
 TRADE DIARY — WINDOWS DESKTOP APP
====================================


 WHAT'S IN THIS FOLDER?
------------------------
  trade_diary.py              → The app (Python source)
  Trade_Diary_Backtest_1.xlsx → Your Excel file (edited automatically)
  Launch_Trade_Diary.bat      → Double-click to run the app instantly
  Build_EXE.ps1               → Builds a standalone Trade_Diary.exe (no Python needed after)
  README.txt                  → This file


 ──────────────────────────────────────────
  OPTION A — Run directly (needs Python)
 ──────────────────────────────────────────
 1. Install Python from https://www.python.org/downloads/
    ⚠ Check "Add Python to PATH" during installation!

 2. Double-click:  Launch_Trade_Diary.bat
    → Installs packages on first run (takes ~30 seconds)
    → Opens the app as a native Windows window

 Every time after: just double-click Launch_Trade_Diary.bat


 ──────────────────────────────────────────
  OPTION B — Build a real .exe (recommended)
 ──────────────────────────────────────────
 Do this once to get a standalone Trade_Diary.exe that
 works on any Windows PC without Python installed.

 1. Install Python (same as Option A, step 1)

 2. Right-click Build_EXE.ps1 → "Run with PowerShell"
    → Takes 1-2 minutes to build

 3. When done, a folder called "dist" opens automatically
    Inside it: Trade_Diary.exe

 4. Copy these two files to any folder you like:
      • Trade_Diary.exe
      • Trade_Diary_Backtest_1.xlsx
    ⚠ They MUST stay in the same folder!

 5. Double-click Trade_Diary.exe to launch


 HOW THE APP WORKS
-------------------
 • Every trade you Add / Edit / Remove is saved DIRECTLY
   to Trade_Diary_Backtest_1.xlsx in the same folder.
 • No import/export needed. No server. No internet.
 • You can open the Excel file in Excel at any time
   (just not while the app is open at the same time).


 TROUBLESHOOTING
-----------------
 "Python not found":
   → Reinstall Python, check "Add to PATH"

 PowerShell says "script cannot be run":
   → Open PowerShell as Admin and run:
      Set-ExecutionPolicy RemoteSigned
   → Then try Build_EXE.ps1 again

 App opens but Excel not found:
   → Make sure Trade_Diary_Backtest_1.xlsx is in
      the same folder as trade_diary.py (or .exe)

====================================
