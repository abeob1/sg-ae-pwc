@ECHO OFF

REM The following directory is for .NET 4.0
set DOTNETFX2=%SystemRoot%\Microsoft.NET\Framework\v4.0.30319
set PATH=%PATH%;%DOTNETFX2%

echo Installing IEPPAMS Win Service...
echo ---------------------------------------------------
C:\Windows\Microsoft.NET\Framework\v4.0.30319\InstallUtil AE_PWC_WS01.exe

echo ---------------------------------------------------
pause
echo Done.