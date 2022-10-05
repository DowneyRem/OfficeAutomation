@echo off
REM echo 切换至 UTF8
chcp 65001 >nul
REM echo 开启延迟变量
setlocal EnableDelayedExpansion
REM test=1 输出测试内容
set test=0


if %test%==1 echo 当前目录
echo %~dp0
cd /d %~dp0
goto FindPython


:FindPython
if %test%==1 echo 寻找 Python 程序路径
dir /s/b | findstr "python.exe" > pythonpath.txt
for /f "usebackq delims=," %%a in (pythonpath.txt) do (
	set "PythonPath=%%~a"
	if %test%==1 echo !PythonPath!
	)
del /q pythonpath.txt
	
REM 循环寻找Python程序，找到后运行脚本
if "%cd%"=="%~d0\" (
	if "!PythonPath!"=="" (
		echo 当前磁盘下没有 python.exe 
		goto End
		)
	)

if "%PythonPath%"=="" (
	if %test%==1 echo 当前目录下没有 python.exe 
	goto BackToParentPath
	)^
else (
	if %test%==1 echo.
	if %test%==1 echo 已经找到 python.exe
	if %test%==1 echo !PythonPath:"=!
	goto RunScript
	)
	

:TraversalDisk
for /L %%a in (67, 1, 90) do (
	cmd /c exit %%a
	set disk=!=exitcodeAscii!:\
	if exist !disk! (
		echo !disk!
		cd !disk!
		goto FindPython
		)
	)


:BackToParentPath
if %test%==1 echo.
if %test%==1 echo 返回上级目录
for %%a in ("%cd%") do set ParentPath=%%~dpa
if %test%==1 echo %ParentPath%
cd %ParentPath%
goto FindPython


:RunScript
if %test%==1 echo.
if %test%==1 echo 寻找 Python 脚本路径
set PythonFile="%~f0"
set PythonFile=%PythonFile:.cmd=.py%
title %PythonFile%
if %test%==1 echo %PythonFile%

if %test%==1 echo.
if %test%==1 echo 执行脚本
echo USE %PythonPath:"=%
echo RUN %PythonFile:"=%
echo.
if %test%==1 (
	call "%PythonPath%"
	)^
else (
	call "%PythonPath%" "%PythonFile%"
	)
goto End


:End
pause & exit