@echo off
chcp 65001 > nul
cd /d "%~dp0"
echo ============================================================
echo   텔레그램 알림 연결 테스트 (Chat ID 자동 탐색 포함)
echo ============================================================
python telegram_notify.py test
echo.
pause
