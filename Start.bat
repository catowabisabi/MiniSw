@echo off
REM Mini Solidwork C# Marco Automator 啟動腳本

echo ========================================
echo Mini Solidwork C# Marco Automator
echo ========================================
echo.

REM 檢查是否存在發布的 EXE
if exist "publish\MiniSolidworkAutomator.exe" (
    echo 正在啟動應用程序...
    echo.
    start "" "publish\MiniSolidworkAutomator.exe"
    exit /b 0
) else (
    echo 錯誤: 未找到已發布的應用程序。
    echo 請先執行以下命令進行發布:
    echo.
    echo   dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o ./publish
    echo.
    pause
    exit /b 1
)
