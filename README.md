# eurobelleza_rpa

# Generar .exe
python -m pip install pyinstaller
Remove-Item -Recurse -Force build -ErrorAction SilentlyContinue
Remove-Item -Recurse -Force dist -ErrorAction SilentlyContinue
Remove-Item -Force Bot.spec -ErrorAction SilentlyContinue
python -m PyInstaller --onefile Bot.py