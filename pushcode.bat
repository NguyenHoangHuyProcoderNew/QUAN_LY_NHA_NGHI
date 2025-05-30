@echo off
REM --- Đổi thành nhánh bạn muốn push ---
set "BRANCH=main"

REM --- Thông điệp commit ---
set "MESSAGE=%~1"

if "%MESSAGE%"=="" (
  echo Vui long nhap thong diep commit.
  echo Usage: pushcode.bat "Thong diep commit"
  goto :eof
)

echo Adding all changes...
git add .

echo Commit voi message: %MESSAGE%
git commit -m "%MESSAGE%"

echo Dang push len remote...
git push origin %BRANCH%

echo Done!
pause
