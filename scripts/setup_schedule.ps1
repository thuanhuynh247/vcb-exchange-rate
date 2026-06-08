# Task Scheduler Setup - VCB Exchange Rate Daily
# Run as Administrator: Right-click PowerShell → "Run as administrator"
# Then: .\setup_schedule.ps1

$taskName = "TyGiaBanking_Daily"
$scriptPath = "D:\Tygia-Tudong\scripts\get_rates.py"
$pythonExe = (Get-Command python -ErrorAction SilentlyContinue).Source

if (-not $pythonExe) {
    Write-Host "❌ Không tìm thấy Python. Hãy đảm bảo Python đã được thêm vào PATH."
    exit 1
}

Write-Host "Python: $pythonExe"
Write-Host "Script: $scriptPath"

# Action: chạy python get_rates.py
$action = New-ScheduledTaskAction `
    -Execute $pythonExe `
    -Argument """$scriptPath""" `
    -WorkingDirectory "D:\Tygia-Tudong\scripts"

# Trigger: 18h00 mỗi ngày
$trigger = New-ScheduledTaskTrigger -Daily -At 6:00PM

# Settings nâng cao
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 15) `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 10)

# Đăng ký task
Register-ScheduledTask `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -TaskName $taskName `
    -Description "Tự động lấy tỷ giá VCB vào 18h00 mỗi ngày. Retry 3 lần nếu thất bại. Chạy lại khi bật máy nếu bỏ lỡ." `
    -Force

Write-Host ""
Write-Host "✅ Đã thiết lập Task Scheduler:"
Write-Host "   📅 Chạy lúc: 18h00 mỗi ngày"
Write-Host "   🔄 Retry: 3 lần, mỗi 10 phút nếu thất bại"
Write-Host "   ⏰ Timeout: 15 phút"
Write-Host "   🔌 StartWhenAvailable: Chạy lại nếu máy tắt lúc 18h"
Write-Host "   🌐 Chỉ chạy khi có mạng"
