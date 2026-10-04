@echo off
chcp 65001 >nul
cd /d %~dp0
if not exist .venv\Scripts\python.exe (
    echo [!] 未找到虚拟环境 .venv，请先执行:
    echo     python -m venv .venv
    echo     .venv\Scripts\pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
    pause
    exit /b 1
)
echo [1/2] 安装/检查 PyInstaller ...
.venv\Scripts\pip install pyinstaller -q -i https://pypi.tuna.tsinghua.edu.cn/simple
echo [2/2] 开始打包（首次需要几分钟，产物在 dist\）...
REM 如嫌单文件启动慢，可把 --onefile 改为 --onedir（产物为一个文件夹）
.venv\Scripts\pyinstaller --noconfirm --clean --onefile --windowed --name 发票填表工具 ^
  --collect-all rapidocr_onnxruntime ^
  --collect-all tkinterdnd2 ^
  --add-data "invoice_filler\builtin_template1.json;invoice_filler" ^
  main.py
echo.
echo 打包完成：dist\发票填表工具.exe
pause
