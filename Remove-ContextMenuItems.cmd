@echo off
:: Removes the OCR context menu entries added by Install-ContextMenu.cmd

setlocal

set BASE=HKCU\Software\Classes\SystemFileAssociations\image\shell

echo Removing OCR context menu entries...

reg delete "%BASE%\OCR" /f
reg delete "%BASE%\OCRChatMode" /f

echo Done.
endlocal
