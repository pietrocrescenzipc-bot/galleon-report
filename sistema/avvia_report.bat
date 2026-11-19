@echo off
REM Lancia la pipeline del report Galleon. Usato dall'Utilita di pianificazione di Windows.
cd /d "%~dp0"
python pipeline.py >> log\esecuzioni.txt 2>&1
