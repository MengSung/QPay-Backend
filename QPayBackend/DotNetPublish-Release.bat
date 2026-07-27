@echo off
setlocal

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Publish-QPayBackend.ps1" -Configuration Release
set "QPAY_PUBLISH_EXIT=%ERRORLEVEL%"

if not "%QPAY_PUBLISH_EXIT%"=="0" echo [QPayPublish] FAILED. Nothing is ready for deployment.
if "%QPAY_PUBLISH_EXIT%"=="0" echo [QPayPublish] SUCCESS.

pause
exit /b %QPAY_PUBLISH_EXIT%
