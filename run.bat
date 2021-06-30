call update.bat
if not exist env call install.bat
env\Scripts\pip install -U -r requirements.txt
env\Scripts\python.exe main.py
