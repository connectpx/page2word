@echo off
setlocal
cd /d "%~dp0"

set "NODE_EXE="
for /f "delims=" %%N in ('where node 2^>nul') do if not defined NODE_EXE set "NODE_EXE=%%N"
if not defined NODE_EXE if exist "%ProgramFiles%\nodejs\node.exe" set "NODE_EXE=%ProgramFiles%\nodejs\node.exe"
if not defined NODE_EXE if exist "%LocalAppData%\Programs\nodejs\node.exe" set "NODE_EXE=%LocalAppData%\Programs\nodejs\node.exe"
if not defined NODE_EXE (
  for /d %%P in ("%TEMP%\page2word-node-*-portable") do (
    for /d %%N in ("%%~fP\node-v*-win-x64") do (
      if not defined NODE_EXE if exist "%%~fN\node.exe" set "NODE_EXE=%%~fN\node.exe"
    )
  )
)

if not defined NODE_EXE (
  start "Page2Word - Node.js required" cmd /k "echo Node.js 20 or newer was not found. Download it from https://nodejs.org/en/download ^& echo install the Windows Installer, then double-click this file again."
  exit /b 1
)

echo Opening the Page2Word server window.
start "Page2Word Writing Assistant Server" cmd /k ""%NODE_EXE%" "%~dp0scripts\dev-server.mjs""
