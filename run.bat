update.bat
if not exist env install.bat
else echo "a"
env\Scripts\python.exe main.py
