call update.bat
if not exist env call install.bat
env\Scripts\python.exe main.py
