# Script para limpiar el inicio del archivo
$content = Get-Content "C:\Users\Soporte\Microsoft-Office-Installer\instalar-gui.ps1"
# Encontrar la línea que tiene #Requires y ponerla al inicio, removiendo cualquier línea anterior
$requiresLine = $content | Where-Object { $_ -match '^#Requires' }
$otherLines = $content | Where-Object { $_ -notmatch '^#Requires' }
# Escribir solo la línea de #Requires y todo lo demás que venga después
$output = @()
$wroteRequires = $false
foreach ($line in $otherLines) {
    if (-not $wroteRequires -and $line -match '^#Requires') {
        $output += $line
        $wroteRequires = $true
    }
}
# Agregar todas las líneas originales excepto las que removimos al inicio
foreach ($line in $otherLines) {
    if ($line -notmatch '^#Requires' -or ($line -match '^#Requires' -and $wroteRequires)) {
        # Skip the original #Requires line since we already added it
        if ($line -match '^#Requires' -and $wroteRequires) { continue }
        $output += $line
    }
}
# Re-escribir asegurando #Requires al inicio
if (-not $wroteRequires) {
    $output = @("#Requires -RunAsAdministrator") + $output
}
Set-Content "C:\Users\Soporte\Microsoft-Office-Installer\instalar-gui.ps1" $output