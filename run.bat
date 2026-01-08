@echo off
if not exist converter.exe (
    echo Converter not found. Building first...
    call build.bat
)
start converter.exe
