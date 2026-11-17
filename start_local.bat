@echo off
setlocal
set "ROOT=%~dp0"
set "ACTION=%~1"
if "%ACTION%"=="" set "ACTION=start"

set "PY="
if exist "%LOCALAPPDATA%\Python\bin\python.exe" set "PY=%LOCALAPPDATA%\Python\bin\python.exe"
if defined PY goto PYTHON_READY

for /f "delims=" %%P in ('where.exe python 2^>nul') do set "PY=%%P"
if defined PY goto PYTHON_READY

for /f "delims=" %%P in ('where.exe py 2^>nul') do if not defined PY set "PY=%%P"
if defined PY goto PYTHON_READY

echo Python was not found. Install Python and add it to PATH.
pause
exit /b 1

:PYTHON_READY
if /I "%ACTION%"=="start" goto START_SERVER

%PY% "%ROOT%idf_material_web.py" %ACTION% %2 %3 %4 %5
if errorlevel 1 goto FAILED
endlocal
exit /b 0

:START_SERVER
echo Starting IDF material web service.
echo Keep this window open and visit http://127.0.0.1:8765
%PY% "%ROOT%idf_material_web.py" serve --host 127.0.0.1 --port 8765
endlocal
exit /b 0

:FAILED
echo.
echo The command failed. Install dependencies with:
echo %PY% -m pip install -r "%ROOT%requirements.txt"
pause
endlocal
exit /b 1
