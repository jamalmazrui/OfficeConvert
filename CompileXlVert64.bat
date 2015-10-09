@echo off
cls
echo Compiling
if exist XlVert64.exe del XlVert64.exe
"c:\Program Files (x86)\AutoIt3\Aut2Exe\Aut2exe_x64.exe" /in XlVert.au3 /out XlVert64.exe /console
if exist XlVert64.exe echo Done
