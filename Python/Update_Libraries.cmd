@ECHO OFF
Color 80
REM This file was generated by 'PyProgramWorkbook'  
REM
REM It updates/installs all required libraries for the MobaLedLib projects.
REM
REM Attention:
REM This program must be started from the arduino libraries directory
REM

CHCP 65001 >NUL
ECHO ************************************
ECHO  Installing the following libraries
ECHO ************************************
ECHO   FastLED
ECHO.
@if exist "%USERPROFILE%\AppData\Local\Temp\MobaLedLib_build\ESP32\includes.cache" del "%USERPROFILE%\AppData\Local\Temp\MobaLedLib_build\ESP32\includes.cache"
"C:\Program Files (x86)\Arduino\arduino_debug.exe" --install-library "FastLED" 2>&1 | find /v " StatusLogger " | find /v " INFO c.a" | find /v " WARN p.a" | find /v " WARN c.a"
ECHO.
ECHO Error %ERRORLEVEL%
IF ERRORLEVEL 1 Goto ErrorMsg

Exit

:ErrorMsg
   COLOR 4F
   ECHO   ****************************************
   ECHO    Da ist was schief gegangen ;-(
   ECHO   ****************************************
   Pause