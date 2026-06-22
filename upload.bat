@echo off
setlocal enabledelayedexpansion

set /p REPO_URL=Paste GitHub repo URL: 
set /p COMMIT_MSG=Commit message: 

if "%REPO_URL%"=="" (
    echo Repo URL cannot be empty.
    pause
    exit /b 1
)

if "%COMMIT_MSG%"=="" (
    echo Commit message cannot be empty.
    pause
    exit /b 1
)

git init
git branch -M main

git remote remove origin 2>nul
git remote add origin %REPO_URL%

git add .
git status --porcelain > status.txt
set /p HASCHANGES=<status.txt
del status.txt >nul 2>&1

if not defined HASCHANGES (
    echo No changes to commit.
) else (
    git commit -m "%COMMIT_MSG%"
    if errorlevel 1 (
        echo Commit failed.
        pause
        exit /b 1
    )
)

git pull --rebase origin main
if errorlevel 1 (
    echo Pull/rebase failed. If there are conflicts, fix them then run:
    echo git add .
    echo git rebase --continue
    pause
    exit /b 1
)

git push -u origin main
if errorlevel 1 (
    echo Push failed.
    pause
    exit /b 1
)

echo Upload complete!
pause