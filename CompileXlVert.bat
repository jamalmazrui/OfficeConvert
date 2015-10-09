@echo off
cls
echo Compiling
if exist XlVert.exe del XlVert.exe
"c:\Program Files (x86)\AutoIt3\Aut2Exe\Aut2exe.exe" /in XlVert.au3 /console
if exist XlVert.exe echo Done
