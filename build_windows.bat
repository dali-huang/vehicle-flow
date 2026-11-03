@echo off
rem Windows 打包: 生成 车辆录入工具.exe (单文件, 无控制台窗口)
rem 需要: Python 3.9+, 先执行 pip install pyinstaller openpyxl xlrd
pyinstaller --noconsole --onefile --name 车辆录入工具 --collect-all openpyxl --hidden-import xlrd --clean --noconfirm webapp.py
echo.
echo 产物在 dist\车辆录入工具.exe
pause
