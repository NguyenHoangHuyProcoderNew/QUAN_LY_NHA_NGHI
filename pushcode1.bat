@echo off
chcp 65001 >nul
setlocal

set "BRANCH=main"

set /p MESSAGE=Nhập thông điệp commit: 

if "%MESSAGE%"=="" (
  echo Bạn chưa nhập thông điệp commit. Thoát chương trình.
  goto :eof
)

echo Tiến hành đẩy toàn bộ dữ liệu trong dự án lên github...
git add .

echo Nội dung commit dự án lên github: %MESSAGE%
git commit -m "%MESSAGE%"

echo Đang cho tải dự án lên github...
git push origin %BRANCH%

echo.
echo Sao lưu dữ liệu lên github thành công!
echo Nhấn phím bất kỳ để thoát... (Tự động thoát sau 5 giây)
timeout /t 5 >nul
exit
