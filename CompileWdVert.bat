@echo off
cls
echo Compiling
if exist WdVert.exe del WdVert.exe
"c:\Program Files (x86)\AutoIt3\Aut2Exe\Aut2exe.exe" /in WdVert.au3 /console
if exist WdVert.exe echo Done
