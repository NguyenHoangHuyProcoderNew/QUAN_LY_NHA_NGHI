@echo off
chcp 65001 >nul
setlocal

set "BRANCH=main"

set /p MESSAGE=Nhập thông điệp commit: 

if "%MESSAGE%"=="" (
  echo Bạn chưa nhập thông điệp commit. Thoát chương trình.
  goto :eof
)

echo Đang thêm thay đổi...
git add .

echo Commit với message: %MESSAGE%
git commit -m "%MESSAGE%"

echo Đang đẩy lên remote...
git push origin %BRANCH%

echo Push code lên GitHub thành công!
<nul set /p= "Nhấn phím bất kỳ để thoát..."
pause >nul
