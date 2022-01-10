@echo off
echo %1

powershell.exe -executionpolicy bypass -file T:\!!!_PowerShell_scripts\MyWiki\mw_main.ps1 %1

rem pause