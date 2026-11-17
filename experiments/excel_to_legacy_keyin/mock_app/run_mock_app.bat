@echo off
rem Launch the practice "legacy" app  (login: admin / 1234)
rem   run_mock_app.bat                   -> starts with login window
rem   run_mock_app.bat -NoLogin          -> skip login
rem   run_mock_app.bat -NoLogin -NoDelay -> skip login and "processing" delay
start "" powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0LegacyOrderApp.ps1" %*
