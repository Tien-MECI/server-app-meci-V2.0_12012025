# encode-credentials.ps1
param (
    [string]$FilePath = "credentials.json"
)

if (-Not (Test-Path $FilePath)) {
    Write-Host "File $FilePath not found!" -ForegroundColor Red
    exit 1
}

# Read file JSON
$bytes = [System.Text.Encoding]::UTF8.GetBytes((Get-Content -Path $FilePath -Raw))

# Convert to Base64
$base64 = [Convert]::ToBase64String($bytes)

# Copy to clipboard
$base64 | Set-Clipboard

Write-Host "Done! Base64 string copied to clipboard."
Write-Host "Now you can paste it (Ctrl+V) into Render -> Env Var GOOGLE_CREDENTIALS_B64"
