@echo off
chcp 65001 >nul
setlocal EnableExtensions EnableDelayedExpansion

REM ==================================================
REM 遊戯王PSA10買取表 更新 → CSV出力 → HTML生成 → GitHubへpush
REM Git管理フォルダも作業フォルダも C:\Users\user\ClimaxGit\climax6 に統一版
REM ==================================================

set "REPO_DIR=C:\Users\user\ClimaxGit\climax6"
set "PYTHON_EXE=python"

REM 既存スクリプト
set "SCRIPT_UPDATE=update_buylist_from_lounge_yugioh_psa10_overwrite_v12.py"
set "SCRIPT_GEN=gen_yugioh_buylist_page_hide_parens_except.py"

REM CSV出力用
set "SCRIPT_CSV=export_buylist_to_csv.py"
set "XLSM_FILE=buylist.xlsm"
set "CSV_FILE=yugioh_psa10_myca_upload.csv"

REM ==================================================
REM S3 / CloudFront deploy settings
REM ==================================================
set "S3_BUCKET=climax-kaitori-static"
set "S3_PREFIX=yugioh"
set "CF_DIST_ID=E51XRDVR8AQAD"
set "PUBLIC_URL=https://kaitori.climax-card.com/yugioh/default/"

cd /d "%REPO_DIR%" || (
  echo [ERROR] フォルダに移動できません: %REPO_DIR%
  pause
  exit /b 1
)

echo.
echo ===== ENV =====
echo REPO_DIR=%REPO_DIR%
echo PYTHON_EXE=%PYTHON_EXE%
echo SCRIPT_UPDATE=%SCRIPT_UPDATE%
echo SCRIPT_CSV=%SCRIPT_CSV%
echo XLSM_FILE=%XLSM_FILE%
echo CSV_FILE=%CSV_FILE%
echo SCRIPT_GEN=%SCRIPT_GEN%
echo ===============
echo.

REM ==================================================
REM 0. Git最新化
REM ==================================================
echo [0/5] git pull --rebase

git pull --rebase
if errorlevel 1 (
  echo [ERROR] git pull --rebase に失敗しました。
  pause
  exit /b 1
)

REM ==================================================
REM 1. 買取表更新 xlsm生成/上書き
REM ==================================================
echo.
echo [1/5] 買取表更新

"%PYTHON_EXE%" "%REPO_DIR%\%SCRIPT_UPDATE%"
if errorlevel 1 (
  echo [ERROR] 買取表更新に失敗しました: %SCRIPT_UPDATE%
  pause
  exit /b 1
)

REM ==================================================
REM 2. buylist.xlsm を CSV に変換
REM ==================================================
echo.
echo [2/5] CSV出力

if not exist "%REPO_DIR%\%XLSM_FILE%" (
  echo [ERROR] 入力ファイルが見つかりません: "%REPO_DIR%\%XLSM_FILE%"
  pause
  exit /b 1
)

"%PYTHON_EXE%" "%REPO_DIR%\%SCRIPT_CSV%" "%REPO_DIR%\%XLSM_FILE%" "%REPO_DIR%\%CSV_FILE%"
if errorlevel 1 (
  echo [ERROR] CSV出力に失敗しました: %SCRIPT_CSV%
  pause
  exit /b 1
)

REM ==================================================
REM 3. HTML生成
REM ==================================================
echo.
echo [3/5] HTML生成

"%PYTHON_EXE%" "%REPO_DIR%\%SCRIPT_GEN%"
if errorlevel 1 (
  echo [ERROR] HTML生成に失敗しました: %SCRIPT_GEN%
  pause
  exit /b 1
)

REM ==================================================
REM 4. Git commit
REM ==================================================
echo.
echo [4/5] git add / commit

git add .
git diff --cached --quiet
if not errorlevel 1 (
  echo [INFO] 変更がないためcommitをスキップします。
) else (
  git commit -m "update yugioh psa10 buylist and csv"
  if errorlevel 1 (
    echo [ERROR] git commit に失敗しました。
    pause
    exit /b 1
  )
)

REM ==================================================
REM 5. Git push
REM ==================================================
echo.
echo [5/5] git push

git push
if errorlevel 1 (
  echo [ERROR] git push に失敗しました。
  pause
  exit /b 1
)

echo.
echo [S3] deploy to kaitori.climax-card.com/%S3_PREFIX%/

where aws >nul 2>&1
if errorlevel 1 (
  echo [ERROR] AWS CLI not found.
  pause
  exit /b 1
)

aws sts get-caller-identity >nul 2>&1
if errorlevel 1 (
  echo [ERROR] AWS credential is invalid or not configured.
  pause
  exit /b 1
)

if not exist "%REPO_DIR%\docs" (
  echo [ERROR] docs folder not found: "%REPO_DIR%\docs"
  pause
  exit /b 1
)

aws s3 sync "%REPO_DIR%\docs" "s3://%S3_BUCKET%/%S3_PREFIX%/" --delete --cache-control "no-cache, no-store, must-revalidate"
if errorlevel 1 (
  echo [ERROR] S3 sync failed.
  pause
  exit /b 1
)

aws cloudfront create-invalidation --distribution-id "%CF_DIST_ID%" --paths "/%S3_PREFIX%/*"
if errorlevel 1 (
  echo [ERROR] CloudFront invalidation failed.
  pause
  exit /b 1
)

echo.
echo [OK] completed.
echo URL: %PUBLIC_URL%
echo CSV: "%REPO_DIR%\%CSV_FILE%"
pause
exit /b 0
