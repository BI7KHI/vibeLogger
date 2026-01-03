@echo off
echo.
echo ===========================================
echo   业余无线电台网日志助手 GUI 版
echo   VibeLogger GUI v2026
echo ===========================================
echo.
echo 正在启动程序...
echo.

REM 检查是否存在可执行文件
if not exist "VibeLogger_GUI.exe" (
    echo 错误：找不到 VibeLogger_GUI.exe
    echo 请确保此批处理文件与 VibeLogger_GUI.exe 在同一目录下
    pause
    exit /b 1
)

REM 启动程序
start "" "VibeLogger_GUI.exe"

echo 程序已启动！
echo.
echo 提示：
echo - 首次运行会自动创建配置文件和Excel文件
echo - 支持GUI表单录入和命令行录入两种模式
echo - 数据自动保存到Excel和CSV文件
echo - 具备智能匹配和自学习功能
echo.
echo 按任意键关闭此窗口...
pause >nul
