@echo off
python -m pip install -r requirements.txt
pyinstaller --noconfirm --onefile --windowed --add-data "prefix_db.json;." phone_tool.py
