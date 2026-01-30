English Word Rotator (Tkinter)

What it does
- Shows one word at a time in a small window
- Advances to the next word at your chosen interval (default 15 seconds)
- Shows a meaning and an optional image for each word
- Lets you add words and save them to words.csv

Files
- english_app.py
- words.csv (optional; word, meaning, image columns)

How to install and run (Windows)
1) Install Python 3
   - Download from python.org and check the box "Add Python to PATH" during install.
2) Open PowerShell
3) Go to the app folder:
   cd C:\Users\Administrator\Documents\Englishapp
4) (Optional) Create words.csv with your words, meanings, and image paths.
5) Run the app:
   python english_app.py

Dependencies
- Tkinter (comes bundled with standard Python on Windows; no extra install needed)

Notes
- If Python is not found, try: py english_app.py
- You can change the interval from inside the app and click "Apply Interval"
- Images should be PNG files for Tkinter's built-in image support

Packaging to a standalone .exe (PyInstaller)
1) Install PyInstaller:
   py -m pip install --upgrade pyinstaller
2) Build the exe:
   py -m PyInstaller --onefile --windowed english_app.py
3) Find the exe at:
   .\dist\english_app.exe
4) Copy english_app.exe to any Windows PC and run it.

Optional (include words.txt next to the exe)
- Copy words.txt into the same folder as english_app.exe

Installer (simple zip distribution)
- Zip the .exe (and optional words.txt) and share the zip.
