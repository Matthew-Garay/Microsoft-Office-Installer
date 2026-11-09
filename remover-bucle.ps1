# Script para remover bloque de comentarios
$content = Get-Content "C:\Users\Soporte\Microsoft-Office-Installer\instalar-gui.ps1"
$filtered = $content | Where-Object { 
    $_ -notmatch '^﻿<#' -and 
    $_ -notmatch '^\.SYNOPSIS' -and 
    $_ -notmatch '^\.DESCRIPTION' -and 
    $_ -notmatch '^Graphical tool' -and 
    $_ -notmatch '^Requirements: Administrator' -and 
    $_ -notmatch '^Usage: irm https://raw' -and 
    $_ -notmatch '^#>$' -and 
    $_ -notmatch '^$'
}
Set-Content "C:\Users\Soporte\Microsoft-Office-Installer\instalar-gui.ps1" $filtered