@echo off
setlocal

set "MANIFEST=%~dp0manifest.xml"
set "DEVKEY=HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"

powershell -NoProfile -Command ^
  "$dev = '%DEVKEY%';" ^
  "$manifest = [System.IO.Path]::GetFullPath($env:MANIFEST);" ^
  "$removed = 0;" ^
  "if (Test-Path $dev) {" ^
  "  $props = (Get-ItemProperty -Path $dev).PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' };" ^
  "  foreach ($p in $props) {" ^
  "    if ([string]$p.Value -ieq $manifest) {" ^
  "      Remove-ItemProperty -Path $dev -Name $p.Name -ErrorAction SilentlyContinue;" ^
  "      $removed++;" ^
  "    }" ^
  "  }" ^
  "}" ^
  "if ($removed -eq 0) { exit 2 } else { exit 0 }"

if errorlevel 2 (
  echo [INFO] Tidak ada entry sideload Skripsi Helper yang cocok untuk dihapus.
) else (
  if errorlevel 1 (
    echo [ERROR] Gagal menghapus entry sideload.
    pause
    exit /b 1
  ) else (
    echo [OK] Entry sideload Skripsi Helper sudah dihapus.
    echo Tutup semua Word, lalu buka Word lagi.
  )
)
pause
exit /b 0
