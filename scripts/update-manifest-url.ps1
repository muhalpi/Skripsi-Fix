param(
  [Parameter(Mandatory = $true)]
  [string]$BaseUrl
)

$manifestPath = Join-Path $PSScriptRoot "..\public\manifest.xml"
$content = Get-Content -LiteralPath $manifestPath -Raw

$content = $content -replace "https://localhost:3000", $BaseUrl

Set-Content -LiteralPath $manifestPath -Value $content -Encoding utf8
Write-Host "Manifest updated to base URL: $BaseUrl"
