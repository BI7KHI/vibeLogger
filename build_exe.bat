@echo off
echo.
echo ===========================================
echo   VibeLogger 打包脚本
echo ===========================================
echo.

REM 检查是否存在源文件
if not exist "VibeLogger_gui.py" (
    echo 错误：找不到 VibeLogger_gui.py
    echo 请在包含源代码的目录运行此脚本
    pause
    exit /b 1
)

echo 正在检查PyInstaller...
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo 安装 PyInstaller...
    python -m pip install pyinstaller
    if errorlevel 1 (
        echo 错误：无法安装 PyInstaller
        pause
        exit /b 1
    )
)

echo.
echo 开始打包...
echo.

REM 删除旧的构建文件
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist
if exist "*.spec" del /q *.spec

REM 打包程序
python -m PyInstaller --onefile --noconsole --name "VibeLogger_GUI" VibeLogger_gui.py

if errorlevel 1 (
    echo.
    echo 打包失败！
    pause
    exit /b 1
)

echo.
echo 打包成功！
echo.
echo 生成的文件：
dir dist\*.exe /B

echo.
echo 正在创建启动脚本...

REM 创建启动脚本
echo @echo off > dist\启动VibeLogger.bat
echo echo. >> dist\启动VibeLogger.bat
echo echo =========================================== >> dist\启动VibeLogger.bat
echo echo   业余无线电台网日志助手 GUI 版 >> dist\启动VibeLogger.bat
echo echo   VibeLogger GUI v2026 >> dist\启动VibeLogger.bat
echo echo =========================================== >> dist\启动VibeLogger.bat
echo echo. >> dist\启动VibeLogger.bat
echo echo 正在启动程序... >> dist\启动VibeLogger.bat
echo if not exist "VibeLogger_GUI.exe" ^( >> dist\启动VibeLogger.bat
echo     echo 错误：找不到可执行文件 >> dist\启动VibeLogger.bat
echo     pause >> dist\启动VibeLogger.bat
echo     exit /b 1 >> dist\启动VibeLogger.bat
echo ^) >> dist\启动VibeLogger.bat
echo start "" "VibeLogger_GUI.exe" >> dist\启动VibeLogger.bat
echo echo 程序已启动！ >> dist\启动VibeLogger.bat

echo.
echo 完成！可执行文件位于 dist\ 目录中
echo.
echo 文件列表：
dir dist\ /B

echo.
echo 按任意键关闭...
pause >nul
