@echo off
echo ========================================
echo 批量Office工具 - 自动编译脚本
echo ========================================
echo.

REM 检查 .NET SDK
dotnet --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到 .NET SDK，请先安装 .NET 6.0 SDK
    echo 下载地址: https://dotnet.microsoft.com/download/dotnet/6.0
    pause
    exit /b 1
)

echo [1/3] 清理旧文件...
rmdir /s /q bin 2>nul
rmdir /s /q obj 2>nul
rmdir /s /q publish 2>nul

echo [2/3] 编译项目...
dotnet publish -c Release -o ./publish -p:PublishSingleFile=true -p:SelfContained=true --runtime win-x64
if errorlevel 1 (
    echo [错误] 编译失败！
    pause
    exit /b 1
)

echo [3/3] 完成！
echo.
echo ========================================
echo ✓ 编译完成！
echo ========================================
echo.
echo exe 文件位置: %cd%\publish\BatchOfficeTool.exe
echo.
pause
