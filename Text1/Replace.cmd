@echo off
chcp 65001 >nul
REM 切换至 UTF8

REM echo 切换回当前文件夹所在路径
REM echo %~dp0
cd /d %~dp0

REM echo 切换回当前文件夹上级路径
REM for %%a in ("%cd%") do set parentpath=%%~dpa
REM echo %parentpath%
REM cd %parentpath%

REM echo 确定 Python 程序路径
dir /s/b | findstr "python.exe" > pythonpath.txt
for /f %%a in (pythonpath.txt) do set PythonPath="%%~a"
del /q pythonpath.txt

REM echo 确定 Python 脚本路径
set PythonFile="%~f0"
set PythonFile=%PythonFile:.cmd=.py%
title %PythonFile%

REM echo 执行脚本
echo USE %PythonPath:"=%
echo RUN %PythonFile:"=%
echo.
REM call %PythonPath%
call %PythonPath% %PythonFile%
pause & exit
