@echo off
rem Drag Packing List file(s) onto this .bat, or double-click it to pick files.
python "%~dp0pl_to_ved.py" %*
if not "%~1"=="" pause
