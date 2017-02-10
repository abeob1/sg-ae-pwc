@ECHO OFF

REM The following directory is for .NET 4.0
set DOTNETFX2=%SystemRoot%\Microsoft.NET\Framework\v4.0.30319
set PATH=%PATH%;%DOTNETFX2%

echo Installing IEPPAMS Win Service...
echo ---------------------------------------------------
C:\Windows\Microsoft.NET\Framework\v4.0.30319\InstallUtil WindowsService1.exe
net start WindowsService1
echo ---------------------------------------------------
pause
echo Done.