@echo off
chcp 65001 >nul
title Ho so khao sat - server
cd /d "%~dp0"

set PORT=8765
set URL=http://127.0.0.1:%PORT%/ho-so-khao-sat.html

set PY=python
where python >nul 2>nul || set PY=py
%PY% --version >nul 2>nul || goto nopython

rem Server dang chay san thi chi mo trinh duyet
netstat -ano | findstr /c:"127.0.0.1:%PORT% " | findstr LISTENING >nul && goto openonly

rem Dong goi lai de nhan mau Word moi sua (loi thi van dung ban cu)
%PY% build.py
set PAGE=ho-so-khao-sat.html
if not exist %PAGE% set PAGE=index.html
if not exist %PAGE% goto nohtml
set URL=http://127.0.0.1:%PORT%/%PAGE%

echo.
echo  Dang chay tai %URL%
echo  Dong cua so nay de tat server.
echo.
if not defined NO_BROWSER start "" /b cmd /c "ping -n 2 127.0.0.1 >nul & start %URL%"
%PY% -m http.server %PORT% --bind 127.0.0.1
pause
exit /b

:openonly
if not defined NO_BROWSER start "" %URL%
exit /b

:nopython
echo Khong tim thay Python. Ban van co the mo truc tiep file ho-so-khao-sat.html
pause
exit /b 1

:nohtml
echo Khong tao duoc trang web - xem loi phia tren.
pause
exit /b 1
