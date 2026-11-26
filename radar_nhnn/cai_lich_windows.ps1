# Radar NHNN – chạy bộ cào trên máy Windows (phương án dự phòng)
# Dùng khi GitHub Actions bị NHNN chặn hoặc anh muốn chạy từ IP Việt Nam.
# Cách dùng (PowerShell, trong thư mục gốc của kho vcb-exchange-rate):
#   powershell -ExecutionPolicy Bypass -File radar_nhnn\cai_lich_windows.ps1
# Gỡ lịch:  schtasks /Delete /TN "RadarNHNN" /F

$ErrorActionPreference = "Stop"
$Repo = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$Runner = Join-Path $PSScriptRoot "chay_mot_lan.ps1"

@"
`$ErrorActionPreference = 'Continue'
Set-Location '$Repo'
`$env:RADAR_RUNNER = 'may-windows'
`$log = Join-Path '$Repo' 'radar_nhnn\lan_chay.log'
"===== `$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') =====" | Out-File -Append -Encoding utf8 `$log
git pull --rebase --autostash 2>&1 | Out-File -Append -Encoding utf8 `$log
python radar_nhnn\crawler.py 2>&1 | Out-File -Append -Encoding utf8 `$log
git add docs/radar/data 2>&1 | Out-File -Append -Encoding utf8 `$log
git diff --staged --quiet
if (`$LASTEXITCODE -ne 0) {
  git commit -m "chore(radar-nhnn): cập nhật số liệu NHNN từ máy Windows" 2>&1 | Out-File -Append -Encoding utf8 `$log
  git push 2>&1 | Out-File -Append -Encoding utf8 `$log
}
"@ | Out-File -Encoding utf8 $Runner

python -m pip install -r (Join-Path $PSScriptRoot "requirements.txt")

# 30 phút/lần từ 07:00 đến 17:30, thứ 2–6
schtasks /Create /F /TN "RadarNHNN" /SC WEEKLY /D MON,TUE,WED,THU,FRI /ST 07:00 /RI 30 /DU 10:30 `
  /TR "powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$Runner`""

Write-Host "Đã cài lịch RadarNHNN. Chạy thử ngay..."
powershell -NoProfile -ExecutionPolicy Bypass -File $Runner
Get-Content (Join-Path $PSScriptRoot "lan_chay.log") -Tail 20
