@echo off
cls
echo Compiling
if exist WdVert64.exe del WdVert64.exe
"c:\Program Files (x86)\AutoIt3\Aut2Exe\Aut2exe_x64.exe" /in WdVert.au3 /out WdVert64.exe /console
if exist WdVert64.exe echo Done
