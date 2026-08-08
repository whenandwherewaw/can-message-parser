@echo off
chcp 65001 >nul
cd /d "%~dp0"
start "" pythonw "CAN报文解析工具.py"
