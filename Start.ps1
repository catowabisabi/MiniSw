# Mini Solidwork C# Marco Automator 啟動腳本

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Mini Solidwork C# Marco Automator" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

$exePath = Join-Path $PSScriptRoot "publish\MiniSolidworkAutomator.exe"

if (Test-Path $exePath) {
    Write-Host "正在啟動應用程序..." -ForegroundColor Green
    Write-Host ""
    Start-Process $exePath
} else {
    Write-Host "錯誤: 未找到已發布的應用程序。" -ForegroundColor Red
    Write-Host ""
    Write-Host "請先執行以下命令進行發布:" -ForegroundColor Yellow
    Write-Host "  dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o ./publish" -ForegroundColor White
    Write-Host ""
    Read-Host "按 Enter 鍵退出"
}
