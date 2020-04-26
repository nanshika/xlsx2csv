@title %~dpnx0
@echo off
set tar1=C:\Users\W4\GoogleDrive\Script\xlsx2csv\dist\xlsx2csv.exe
set tar0=C:\Users\W4\GoogleDrive\Script\xlsx2csv\xlsx2csv.py
call :t pygame
call :t scipy
call :t sklearn
call :t torch
call :t Crypto
call :t PIL
call :t PyQt5
call :t _geoip_geolite2
exit

call :t notebook
call :t matplotlib
call :t tensorflow
call :t numpy

call :t boto
call :t botocore
call :t caffe2
call :t cv2
call :t cython
call :t gensim
call :t janome


:t
del /Q C:\Users\W4\GoogleDrive\Script\xlsx2csv\*.spec
rd /S /Q C:\Users\W4\GoogleDrive\Script\xlsx2csv\build\
rd /S /Q C:\Users\W4\GoogleDrive\Script\xlsx2csv\__pycache__\

pyinstaller --noconsole --onefile --exclude %1 %tar0% 2> NUL
ren %tar1% exclude_%1.exe
exit /b

