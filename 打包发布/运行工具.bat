@echo off
cd /d "%~dp0"
echo Starting the tool...
python "发货单对账工具.py"
echo.
echo Done! Press any key to exit...
pause >nul
