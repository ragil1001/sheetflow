@echo off

python -m pip install camoufox openpyxl
if errorlevel 1 exit /b 1
python -m camoufox fetch
if errorlevel 1 exit /b 1

echo.
echo INSTALL SELESAI
pause
