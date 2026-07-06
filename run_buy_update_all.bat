@echo off
chcp 65001 >nul
setlocal EnableExtensions EnableDelayedExpansion

set "EXIT_CODE=0"

REM ==================================================
REM 遊戯王 PSA10 buylist update
REM 置き場所例: C:\ClimaxJobs\climax6\run_buy_update_all.bat
REM ==================================================

set "REPO_DIR=%~dp0"
set "REPO_DIR=%REPO_DIR:~0,-1%"
set "PYTHON_EXE=python"

set "SCRIPT_UPDATE=update_buylist_from_lounge_yugioh_psa10_overwrite_v12.py"
set "SCRIPT_CSV=export_yugioh_psa10_myca_csv.py"
set "SCRIPT_GEN=gen_yugioh_buylist_page_hide_parens_except_holo_relief.py"

set "XLSM_FILE=buylist.xlsm"
set "CSV_FILE=yugioh_psa10_myca_upload.csv"

set "S3_BUCKET=climax-kaitori-static"
set "S3_PREFIX=yugioh"
set "CF_DIST_ID=E51XRDVR8AQAD"
set "PUBLIC_URL=https://kaitori.climax-card.com/yugioh/default/"

cd /d "%REPO_DIR%" || (
  echo [ERROR] Cannot cd to: %REPO_DIR%
  set "EXIT_CODE=1"
  goto :END
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
echo S3=s3://%S3_BUCKET%/%S3_PREFIX%/
echo URL=%PUBLIC_URL%
echo ===============
echo.

REM ==================================================
REM 0. Optional git pull. Skip when working tree is dirty.
REM ==================================================
echo [0/6] optional git pull

git rev-parse --is-inside-work-tree >nul 2>&1
if errorlevel 1 (
  echo [WARN] This folder is not a git repository. Continue without git.
  goto :RUN_UPDATE
)

git diff --quiet
set "DIRTY1=!errorlevel!"
git diff --cached --quiet
set "DIRTY2=!errorlevel!"

if not "!DIRTY1!!DIRTY2!"=="00" (
  echo [WARN] Local changes exist. Skip initial git pull to avoid rebase error.
) else (
  git pull --rebase
  if errorlevel 1 (
    echo [WARN] git pull failed. Continue update and publish anyway.
  )
)

:RUN_UPDATE
REM ==================================================
REM 1. Update Excel buylist
REM ==================================================
echo.
echo [1/6] update buylist

if not exist "%REPO_DIR%\%SCRIPT_UPDATE%" (
  echo [ERROR] Missing update script: %REPO_DIR%\%SCRIPT_UPDATE%
  set "EXIT_CODE=1"
  goto :END
)

"%PYTHON_EXE%" "%REPO_DIR%\%SCRIPT_UPDATE%"
if errorlevel 1 (
  echo [ERROR] Update failed: %SCRIPT_UPDATE%
  set "EXIT_CODE=1"
  goto :END
)

REM ==================================================
REM 2. Export CSV
REM ==================================================
echo.
echo [2/6] export CSV

if not exist "%REPO_DIR%\%XLSM_FILE%" (
  echo [ERROR] Missing Excel file: %REPO_DIR%\%XLSM_FILE%
  set "EXIT_CODE=1"
  goto :END
)

if not exist "%REPO_DIR%\%SCRIPT_CSV%" (
  echo [ERROR] Missing CSV script: %REPO_DIR%\%SCRIPT_CSV%
  echo [HINT] Check this folder with: dir /b *csv*.py
  set "EXIT_CODE=1"
  goto :END
)

"%PYTHON_EXE%" "%REPO_DIR%\%SCRIPT_CSV%" "%REPO_DIR%\%XLSM_FILE%" "%REPO_DIR%\%CSV_FILE%"
if errorlevel 1 (
  echo [ERROR] CSV export failed: %SCRIPT_CSV%
  set "EXIT_CODE=1"
  goto :END
)

REM ==================================================
REM 3. Generate HTML
REM ==================================================
echo.
echo [3/6] generate HTML

if not exist "%REPO_DIR%\%SCRIPT_GEN%" (
  echo [ERROR] Missing HTML script: %REPO_DIR%\%SCRIPT_GEN%
  set "EXIT_CODE=1"
  goto :END
)

"%PYTHON_EXE%" "%REPO_DIR%\%SCRIPT_GEN%"
if errorlevel 1 (
  echo [ERROR] HTML generation failed: %SCRIPT_GEN%
  set "EXIT_CODE=1"
  goto :END
)

REM ==================================================
REM 4. Git commit/push. Do not stop S3 publish on git errors.
REM ==================================================
echo.
echo [4/6] git commit and push best-effort

git rev-parse --is-inside-work-tree >nul 2>&1
if errorlevel 1 (
  echo [WARN] Not a git repository. Skip git commit/push.
  goto :S3_DEPLOY
)

git add .
git diff --cached --quiet
if not errorlevel 1 (
  echo [INFO] No git changes to commit.
) else (
  git commit -m "update yugioh psa10 buylist"
  if errorlevel 1 (
    echo [WARN] git commit failed. Continue to S3 publish.
  )
)

git pull --rebase
if errorlevel 1 (
  echo [WARN] git pull after commit failed. Continue to S3 publish.
) else (
  git push
  if errorlevel 1 (
    echo [WARN] git push failed. Continue to S3 publish.
  )
)

:S3_DEPLOY
REM ==================================================
REM 5. S3 deploy
REM ==================================================
echo.
echo [5/6] S3 deploy

where aws >nul 2>&1
if errorlevel 1 (
  echo [ERROR] AWS CLI not found.
  set "EXIT_CODE=1"
  goto :END
)

aws sts get-caller-identity >nul 2>&1
if errorlevel 1 (
  echo [ERROR] AWS credential is invalid or not configured.
  set "EXIT_CODE=1"
  goto :END
)

if not exist "%REPO_DIR%\docs" (
  echo [ERROR] docs folder not found: %REPO_DIR%\docs
  set "EXIT_CODE=1"
  goto :END
)

aws s3 sync "%REPO_DIR%\docs" "s3://%S3_BUCKET%/%S3_PREFIX%/" ^
  --delete ^
  --cache-control "no-cache, no-store, must-revalidate"

if errorlevel 1 (
  echo [ERROR] S3 sync failed.
  set "EXIT_CODE=1"
  goto :END
)

REM ==================================================
REM 6. CloudFront invalidation
REM ==================================================
echo.
echo [6/6] CloudFront invalidation

aws cloudfront create-invalidation --distribution-id "%CF_DIST_ID%" --paths "/%S3_PREFIX%/*"
if errorlevel 1 (
  echo [ERROR] CloudFront invalidation failed.
  set "EXIT_CODE=1"
  goto :END
)

echo.
echo [OK] completed.
echo URL: %PUBLIC_URL%
echo CSV: %REPO_DIR%\%CSV_FILE%

:END
echo.
echo ==== Finished ====

if "%AUTO_RUN%"=="1" (
  endlocal
  exit /b %EXIT_CODE%
)

pause
endlocal
exit /b %EXIT_CODE%
