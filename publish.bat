@echo off
rem 画面モニター 公開スクリプト（Windows 用）。ダブルクリックで実行。
rem 中身は publish.sh を Git for Windows の bash で動かす。
chcp 65001 >nul
set "BASH=%ProgramFiles%\Git\bin\bash.exe"
if not exist "%BASH%" set "BASH=%LocalAppData%\Programs\Git\bin\bash.exe"
if not exist "%BASH%" set "BASH=%ProgramFiles(x86)%\Git\bin\bash.exe"
if not exist "%BASH%" (
  echo Git for Windows が見つかりません。https://git-scm.com/ からインストールしてください。
  pause
  exit /b 1
)
"%BASH%" "%~dp0publish.sh" %*
echo.
pause
