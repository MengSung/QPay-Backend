@echo off
setlocal EnableExtensions

rem Fast, deployment-ready Release publish for the .NET Framework 4.7.1 web app.
rem Restore once when needed, then publish without repeating clean/restore work.
set "PROJECT=%~dp0QPayBackend.csproj"
set "OUTPUT=%~dp0..\artifacts\QPayBackend-Release"

if not exist "%PROJECT%" (
  echo [QPayPublish] ERROR: project not found: "%PROJECT%"
  exit /b 2
)

if not exist "%~dp0obj\project.assets.json" (
  echo [QPayPublish] Restoring NuGet packages...
  dotnet restore "%PROJECT%" --nologo
  if errorlevel 1 exit /b %errorlevel%
)

echo [QPayPublish] Publishing Release output...
dotnet publish "%PROJECT%" -c Release -f net471 -o "%OUTPUT%" --no-restore --nologo
if errorlevel 1 (
  echo [QPayPublish] FAILED.
  exit /b %errorlevel%
)

echo [QPayPublish] SUCCESS: "%OUTPUT%"
exit /b 0
