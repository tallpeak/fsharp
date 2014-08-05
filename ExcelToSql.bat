@echo off
:An example of calling the fsx script from Fsharp interactive
:fsi.exe must exist in the path or one of the following
set fs=C:\Program Files (x86)\Microsoft SDKs\F#\3.1\Framework\v4.0\Fsi.exe
if not exist "%fs%" set fs=C:\Program Files (x86)\Microsoft SDKs\F#\3.0\Framework\v4.0\Fsi.exe
if not exist "%fs%" set fs=C:\Program Files (x86)\Microsoft F#\v4.0\fsi.exe
if not exist "%fs%" set fs=fsi.exe
if exist "%fs%" goto good
where %fs% > nul 2>&1
if errorlevel 1 goto error
:good
"%fs%" %~dp0\excelToSQL.fs %*
goto end
:error
echo Error: Fsharp interactive (fsi.exe) not found! fs=%fs%
:end
