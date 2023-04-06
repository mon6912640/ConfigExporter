@echo off
set CUR=%~dp0
pyinstaller -F ConfigExporter.py --distpath=%CUR%
pause