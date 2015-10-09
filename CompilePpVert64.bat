@echo off
cls
echo Compiling
if exist PpVert64.exe del PpVert64.exe
"c:\Program Files (x86)\AutoIt3\Aut2Exe\Aut2exe_x64.exe" /in PpVert.au3 /out PpVert64.exe /console
if exist PpVert64.exe echo Done
