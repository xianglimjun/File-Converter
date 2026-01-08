@echo off
set "INPUT_FILE=%~1"
set "OUTPUT_FILE=%~2"
set "EXT=%~x1"
set "LOG=conversion_debug.txt"

echo [DEBUG] %DATE% %TIME% > "%LOG%"
echo [DEBUG] Input: "%INPUT_FILE%" >> "%LOG%"
echo [DEBUG] Output: "%OUTPUT_FILE%" >> "%LOG%"

if "%INPUT_FILE%"=="" (
    echo [ERROR] No input file provided >> "%LOG%"
    goto usage
)

rem Resolve absolute paths
for %%f in ("%INPUT_FILE%") do set "ABS_INPUT=%%~ff"
for %%f in ("%OUTPUT_FILE%") do set "ABS_OUTPUT=%%~ff"

echo [DEBUG] Absolute Input: "%ABS_INPUT%" >> "%LOG%"
echo [DEBUG] Absolute Output: "%ABS_OUTPUT%" >> "%LOG%"

if /i "%EXT%"==".docx" goto word
if /i "%EXT%"==".doc" goto word
if /i "%EXT%"==".pptx" goto ppt
if /i "%EXT%"==".ppt" goto ppt

echo [ERROR] Unsupported file type: %EXT% >> "%LOG%"
exit /b 1

:word
echo [DEBUG] Starting Word conversion... >> "%LOG%"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$w=New-Object -ComObject Word.Application; $w.Visible=$false; $w.DisplayAlerts=0; try { Write-Host 'Opening document...'; $p = $w.Documents.Open('%ABS_INPUT%'); Write-Host 'Saving as PDF...'; $p.SaveAs('%ABS_OUTPUT%', 17); $p.Close(); Write-Host 'Success'; } catch { Write-Host 'Error:' $_; exit 1 } finally { if($w) { $w.Quit() } }" >> "%LOG%" 2>&1
if %errorlevel% neq 0 ( 
    echo [ERROR] Word conversion failed with code %errorlevel% >> "%LOG%"
    exit /b 1 
)
goto end

:ppt
echo [DEBUG] Starting PowerPoint conversion... >> "%LOG%"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$a=New-Object -ComObject PowerPoint.Application; $a.DisplayAlerts=1; try { Write-Host 'Opening presentation...'; $p = $a.Presentations.Open('%ABS_INPUT%', 0, 0, 0); Write-Host 'Saving as PDF...'; $p.SaveAs('%ABS_OUTPUT%', 32); $p.Close(); Write-Host 'Success'; } catch { Write-Host 'Error:' $_; exit 1 } finally { if($a) { $a.Quit() } }" >> "%LOG%" 2>&1
if %errorlevel% neq 0 ( 
    echo [ERROR] PowerPoint conversion failed with code %errorlevel% >> "%LOG%"
    exit /b 1 
)
goto end

:usage
echo Usage: convert_doc.bat "input_file" "output_file" >> "%LOG%"
exit /b 1

:end
echo [DEBUG] Conversion successful. >> "%LOG%"
exit /b 0
