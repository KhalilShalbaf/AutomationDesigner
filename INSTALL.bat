@echo off
setlocal
set TARGET=%LOCALAPPDATA%\AutomationDesigner
if not exist "%TARGET%" mkdir "%TARGET%"
xcopy /E /Y /I "%~dp0" "%TARGET%"
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\AutomationDesigner" /v "Description" /d "AutomationDesigner" /f
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\AutomationDesigner" /v "LoadBehavior" /t REG_DWORD /d 3 /f
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\AutomationDesigner" /v "Manifest" /d "%TARGET%\AutomationDesigner.vsto|vstolocal" /f
echo.
echo Installed to %TARGET%
pause
