@echo off
cd /d "C:\Users\Ewan\OneDrive - Studio Snaidero Chicago\Shared\111W_Logistics Site"
python 111W_update_dashboard.py %*
if errorlevel 1 (
  echo.
  echo *** The update did NOT publish. See the message above and 111W_update_log.txt ***
  echo.
)
pause
