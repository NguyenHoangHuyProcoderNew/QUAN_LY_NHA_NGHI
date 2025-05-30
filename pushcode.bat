@echo off
setlocal

set "BRANCH=main"

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

echo Push Code Len Github Thanh Cong !
<nul set /p= "Nhan phim bat ky de thoat..."
pause >nul
