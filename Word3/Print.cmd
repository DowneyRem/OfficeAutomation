@echo off
chcp 65001 >nul
rem 切换成UTF8

rem 测试：Windows10
rem 理论支持：Win7 及以上
rem Python 3.8.10 支持win7
rem 部分空行用于更好地解析代码，请勿随意删除


set pypath="C:\Program Files\Python310\python.exe"
echo %pypath%
echo.
%pypath% Print.py
pause &exit