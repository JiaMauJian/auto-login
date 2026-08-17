@echo off
rem 雙擊就開持股同步介面。cd /d 是必要的：從別的資料夾雙擊時，
rem 工作目錄不會自動跟著 .bat 走，.env 與 Excel 都會找不到。
cd /d "%~dp0"
python ui.py
if errorlevel 1 pause
