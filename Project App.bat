@echo off
echo Starting Flask Development Server...
cd /D "C:\Users\ACER\Desktop\HRD Portal & Projects\Project_Portal"

REM
if exist venv\Scripts\activate (
    start /B "" cmd /c "call venv\Scripts\activate && flask run --host=0.0.0.0 --port=5000"
) else (
    start /B "" cmd /c "flask run --host=0.0.0.0 --port=5000"
)

echo Waiting for server to start...
timeout /t 3 /nobreak

echo Opening Project Portal in Chrome...
start chrome http://172.23.13.115:5000/