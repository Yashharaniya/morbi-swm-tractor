@echo off
chcp 65001 >nul
title Morbi Municipal Corporation - SWM Tractor Portal
cls
echo =========================================================================
echo    🏛️ મોરબી મહાનગરપાલિકા - સોલિડ વેસ્ટ મેનેજમેન્ટ (SWM)
echo    દૈનિક ટ્રેક્ટર ફેરા મોનિટરિંગ & DyMC એક્સપોર્ટ પોર્ટલ (વર્ઝન ૨.૦)
echo =========================================================================
echo.
echo [1] Checking Python environment and database...
echo.

:: Get Local IP Address for Mobile Access
for /f "tokens=2 delims=:" %%i in ('ipconfig ^| findstr /c:"IPv4 Address" /c:"IPv4" 2^>nul') do (
    set LOCAL_IP=%%i
    goto :ip_found
)
:ip_found
set LOCAL_IP=%LOCAL_IP: =%

echo =========================================================================
echo  🌐 LOCAL PC LINK (કમ્પ્યુટર માટે):
echo     http://localhost:8000
echo.
if not "%LOCAL_IP%"=="" (
    echo  📱 MOBILE WI-FI LINK (SSI અધિકારીઓના ફોન માટે):
    echo     http://%LOCAL_IP%:8000
    echo.
    echo  (તમારા ફોનને આ જ Wi-Fi કે મોબાઈલ હોટસ્પોટ સાથે કનેક્ટ કરી લિંક ખોલો)
)
echo =========================================================================
echo  🔐 SECURITY & PASSWORDS:
echo     - Morbi Logo Click: Master Admin Password = 87654321 (Hint: 81)
echo     - Ward SSI Save Password = 1234 (Editable by Master Admin)
echo     - Render.com 0-Data-Loss: Auto-backup active at data_backup/swm_backup.json
echo =========================================================================
echo.
echo Starting application...
start "" http://localhost:8000
python app.py
pause
