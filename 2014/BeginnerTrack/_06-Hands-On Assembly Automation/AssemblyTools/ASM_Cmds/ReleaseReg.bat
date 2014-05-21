@echo off

set ADDIN_PATH="%~dp0\ASM_Cmds\bin\Release\SolidEdge.ASM_Edgebar_Cmds.dll"
set REGASM_X86="C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
set REGASM_X64="C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"

CLS


echo %ADDIN_PATH%
echo 

echo This batch file must be executed with administrator privileges!
echo. 

:menu
echo [Options]
echo 1 Register (Solid Edge x64)
echo 2 Unregister (Solid Edge x64)
echo 3 Quit

:choice
set /P C=Enter selection:
if "%C%"=="1" goto registerx64
if "%C%"=="2" goto unregisterx64
if "%C%"=="3" goto end
goto choice


:registerx64
set REGASM_PATH=%REGASM_X64%
goto register

:unregisterx64
set REGASM_PATH=%REGASM_X64%
goto unregister

:register
echo.
%REGASM_PATH% /codebase %ADDIN_PATH%
goto end

:unregister
echo.
%REGASM_PATH% /u %ADDIN_PATH%
goto end

:end
pause