# Convertir a UTF-8 sin BOM y cambiar CRLF a LF
$filePath = "C:\Users\Soporte\Microsoft-Office-Installer\instalar-gui.ps1"
$content = Get-Content -Path $filePath
# Unir con LF y sin BOM
$contentWithLF = ($content -join "`n")
# Escribir sin BOM (usando [System.IO.File]::WriteAllLines para evitar BOM)
[System.IO.File]::WriteAllLines($filePath, $contentWithLF, [System.Text.Encoding]::UTF8.GetBytes("test")) > $null
# Alternative approach: write with explicit UTF8 no BOM
$bytes = [System.Text.Encoding]::UTF8.GetBytes($contentWithLF)
# Remove BOM if present (first 3 bytes are 0xEF,0xBB,0xBF)
if ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
    $bytes = $bytes | Where-Object { $_ -ne 0xEF -and $_ -ne 0xBB -and $_ -ne 0xBF }
}
[System.IO.File]::WriteAllBytes($filePath, $bytes)
Write-Host "Conversion completed"