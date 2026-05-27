@echo off
:: Adds OCR and OCR (Chat Mode) to the right-click menu for all image files.
:: %~dp0 resolves to the directory containing this script at install time,
:: so the .cmd file must live next to Get-OCRTextFromGPT.ps1.

setlocal

set PS51=C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
set SCRIPT=%~dp0Get-OCRTextFromGPT.ps1
set BASE=HKCU\Software\Classes\SystemFileAssociations\image\shell

echo Installing OCR context menu entries...
echo Script path: %SCRIPT%

reg add "%BASE%\OCR" /ve /t REG_SZ /d "OCR" /f
reg add "%BASE%\OCR" /v "Icon" /t REG_SZ /d "%PS51%" /f
reg add "%BASE%\OCR\command" /ve /t REG_SZ /d "\"%PS51%\" -NoProfile -File \"%SCRIPT%\" \"%%1\" -ToClipboard" /f

reg add "%BASE%\OCRChatMode" /ve /t REG_SZ /d "OCR (Chat Mode)" /f
reg add "%BASE%\OCRChatMode" /v "Icon" /t REG_SZ /d "%PS51%" /f
reg add "%BASE%\OCRChatMode\command" /ve /t REG_SZ /d "\"%PS51%\" -NoProfile -File \"%SCRIPT%\" \"%%1\" -ToClipboard -ChatMode" /f

echo Done.
endlocal
