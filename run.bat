call update.bat
if not exist env call install.bat
else echo "a"
env\Scripts\python.exe main.py
