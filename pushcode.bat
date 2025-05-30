@echo off
setlocal

REM --- Đổi thành nhánh bạn muốn push ---
set "BRANCH=main"

REM --- Hỏi người dùng nhập commit message ---
set /p MESSAGE=Nhap thong diep commit: 

if "%MESSAGE%"=="" (
  echo Ban chua nhap thong diep commit. Thoat chuong trinh.
  goto :eof
)

echo Adding all changes...
git add .

echo Commit voi message: %MESSAGE%
git commit -m "%MESSAGE%"

echo Dang push len remote...
git push origin %BRANCH%

echo Thanh Cong!
pause
