@echo off
chcp 65001 >nul
rem Windows 打包: 生成 vehicle-flow-windows.exe (单文件, 无控制台窗口)
rem 需要: Python 3.9+, 先执行 pip install pyinstaller openpyxl xlrd
pyinstaller --noconsole --onefile --name vehicle-flow-windows --icon assets\icon.ico --collect-all openpyxl --hidden-import xlrd --clean --noconfirm webapp.py
echo.
echo 产物在 dist\vehicle-flow-windows.exe
pause
