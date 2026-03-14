@echo off
cd /d "%~dp0"
python bank_reconciliation_table.py -i reconciliation_input --gui
pause
