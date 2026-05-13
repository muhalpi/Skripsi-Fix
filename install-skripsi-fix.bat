@echo off
setlocal

set "MANIFEST=%~dp0manifest.xml"
if not exist "%MANIFEST%" (
  echo [ERROR] manifest.xml tidak ditemukan di folder ini.
  echo Pastikan file manifest.xml berada di folder yang sama dengan script ini.
  pause
  exit /b 1
)

set "DEVKEY=HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
powershell -NoProfile -Command ^
  "$dev = '%DEVKEY%';" ^
  "$manifest = [System.IO.Path]::GetFullPath($env:MANIFEST);" ^
  "New-Item -Path $dev -Force | Out-Null;" ^
  "$props = (Get-ItemProperty -Path $dev).PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' };" ^
  "foreach ($p in $props) { if ([string]$p.Value -ieq $manifest) { Remove-ItemProperty -Path $dev -Name $p.Name -ErrorAction SilentlyContinue } };" ^
  "$newGuid = [guid]::NewGuid().ToString();" ^
  "Set-ItemProperty -Path $dev -Name $newGuid -Value $manifest;"

if errorlevel 1 (
  echo [ERROR] Gagal menulis registry sideload.
  pause
  exit /b 1
)

echo [OK] Skripsi-Fix berhasil didaftarkan.
echo Tutup semua Word, lalu buka Word lagi.
echo Jika add-in belum muncul, restart Windows sekali.
pause
exit /b 0
