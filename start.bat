@echo off

rem 定位到脚本所在目录再启动: 从其它目录 cmd /c 调用本文件时也能找到 config_gui.py
cd /d "%~dp0"
pythonw "config_gui.py"
exit
