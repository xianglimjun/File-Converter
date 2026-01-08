@echo off
setlocal

rem Check if cl is already available
where cl >nul 2>nul
if %errorlevel% equ 0 goto build

rem Try to find vcvarsall.bat
set "VS_PATH="
if exist "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\VC\Auxiliary\Build\vcvarsall.bat" set "VS_PATH=C:\Program Files\Microsoft Visual Studio\2022\Enterprise\VC\Auxiliary\Build\vcvarsall.bat"
if exist "C:\Program Files\Microsoft Visual Studio\2022\Professional\VC\Auxiliary\Build\vcvarsall.bat" set "VS_PATH=C:\Program Files\Microsoft Visual Studio\2022\Professional\VC\Auxiliary\Build\vcvarsall.bat"
if exist "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvarsall.bat" set "VS_PATH=C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvarsall.bat"
if exist "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\VC\Auxiliary\Build\vcvarsall.bat" set "VS_PATH=C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\VC\Auxiliary\Build\vcvarsall.bat"

if "%VS_PATH%"=="" (
    echo Error: Could not find Visual Studio installation.
    echo Please run this script from a Developer Command Prompt.
    pause
    exit /b 1
)

echo Found Visual Studio environment at: "%VS_PATH%"
call "%VS_PATH%" x64

:build
echo Building...
rem /MT for static linking (portability), /O2 for optimization
cl /nologo /MT /O2 main.c /Fe:converter.exe user32.lib gdi32.lib comdlg32.lib shell32.lib advapi32.lib
if %errorlevel% neq 0 (
    echo Build failed!
    pause
    exit /b 1
)
echo Build successful!
