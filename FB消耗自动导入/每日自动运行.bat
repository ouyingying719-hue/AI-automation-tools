@echo off
chcp 936 >nul
cd /d "%~dp0"
python -u run_all_v2.py
