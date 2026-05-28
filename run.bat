@echo off
REM Run docconvert on Windows. Activates the thesis-convert conda env
REM and ensures MiKTeX / Strawberry Perl are on PATH.
setlocal
cd /d "%~dp0"

REM --- PATH: TeX distribution + Perl (latexpand) ---
if exist "C:\Users\%USERNAME%\AppData\Local\Programs\MiKTeX\miktex\bin\x64" (
    set "PATH=%PATH%;C:\Users\%USERNAME%\AppData\Local\Programs\MiKTeX\miktex\bin\x64"
)
if exist "C:\Users\%USERNAME%\Strawberry\perl\bin" (
    set "PATH=%PATH%;C:\Users\%USERNAME%\Strawberry\perl\bin"
)

REM --- activate conda env ---
call conda activate thesis-convert || (
    echo Failed to activate conda env "thesis-convert".
    echo Create it with: conda create -n thesis-convert python=3.11 ^&^& conda activate thesis-convert ^&^& pip install -e .
    exit /b 1
)

python convert.py %*
set EXIT_CODE=%ERRORLEVEL%
pause
exit /b %EXIT_CODE%
