@echo off
cls
echo Compiling
if exist PpVert.exe del PpVert.exe
"c:\Program Files (x86)\AutoIt3\Aut2Exe\Aut2exe.exe" /in PpVert.au3 /console
if exist PpVert.exe echo Done
