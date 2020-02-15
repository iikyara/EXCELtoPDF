call update.bat
if not exist env call install.bat
env\Scripts\pip install -r requirements.txt
env\Scripts\python.exe main.py
