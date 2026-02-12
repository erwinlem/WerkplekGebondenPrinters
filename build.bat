@echo off
for /f "tokens=*" %%i in ('git describe') do set VERSION=%%i
rem echo "building WerkplekGebondenPrinter-%VERSION%.zip"
msbuild WerkplekGebondenPrinter.csproj /t:build /p:Version=%VERSION:~1% /p:Configuration=Release /p:DebugSymbols=false /p:DebugType=None /p:OutputPath=build\WerkplekGebondenPrinter-%VERSION%
rem powershell -Command "Compress-Archive -Path build\WerkplekGebondenPrinter-%VERSION%\* -DestinationPath build\WerkplekGebondenPrinter-%VERSION%.zip"