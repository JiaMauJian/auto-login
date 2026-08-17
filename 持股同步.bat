@echo off
rem Opens the stock-sync UI. cd /d is required: double-clicking from
rem Explorer does not make the working directory follow the .bat file,
rem so .env and the Excel file would not be found otherwise.
rem
rem Kept ASCII-only on purpose: cmd.exe parses .bat files using the
rem system's OEM codepage, which is not guaranteed to be UTF-8. Chinese
rem text in a rem comment has been seen to make cmd.exe misread part of
rem the comment as a command on some codepage configurations.
cd /d "%~dp0"
python ui.py
if errorlevel 1 pause
