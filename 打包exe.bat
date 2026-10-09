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
echo [2/2] 开始打包（首次需要几分钟，产物在 dist_安装包\）...
REM 打包参数统一写在 发票填表工具.spec 里；
REM 版本号自动取自 invoice_filler\__init__.py，产物名形如 发票填表工具-v1.0.2.exe
.venv\Scripts\pyinstaller --noconfirm --clean --distpath "dist_安装包" --workpath "build_构建缓存" "发票填表工具.spec"
echo.
echo 打包完成，产物在 dist_安装包\ 目录（文件名带版本号）
pause
