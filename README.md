# screenpop

#Install:
- python -m venv .venv
- Windows: .\.venv\Scripts\activate
- macOS/Linux: source .venv/bin/activate
- python -m pip install --upgrade pip
- python -m pip install flask waitress requests pystray pillow
- Optional (Windows only, for no-activate focus behavior)
- python -m pip install pywin32

#Run
- python screenpop_router.py --port 5588
