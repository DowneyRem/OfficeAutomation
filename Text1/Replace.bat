@echo off
REM echo 切换至 UTF8
chcp 65001 > nul
REM echo 开启延迟变量
setlocal EnableDelayedExpansion
REM test=1 输出测试内容
set test=0


REM 切换文件所在目录
if %test%==1 echo 当前目录
set Folder=%~dp0
echo %Folder%
cd /d %Folder%


REM 设置临时文件路径
set PythonPathTxt="%Folder%PythonPath.txt"
if %test%==1 echo %PythonPathTxt:"=%
echo.
goto FindPython


REM 寻找 Python 程序路径
:FindPython
if %test%==1 echo 寻找 Python 程序路径
dir /s/b | findstr "python.exe" >> %PythonPathTxt%
for /f "usebackq delims=," %%a in (%PythonPathTxt%) do (
	REM set "PythonPath=%%~a"
	set PythonPath=%%~a
	if %test%==1 echo !PythonPath!
	)


if %test%==1 (
	echo. > Nul
	REM del /q %PythonPathTxt%
	)^
else (
	del /q %PythonPathTxt%
	)
	
	
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
	if %test%==1 (
		echo.
		echo 已经找到 python.exe
		echo !PythonPath:"=!
		)
	goto RunScript
	)
	
	
REM 返回上级目录
:BackToParentPath
if %test%==1 (
	echo.
	echo 返回上级目录
	)
for %%a in ("%cd%") do set ParentPath=%%~dpa
if %test%==1 echo %ParentPath%
cd %ParentPath%
goto FindPython


REM 拼接路径，执行脚本
:RunScript
cd /d %Folder%
if %test%==1 (
	echo.
	echo 拼接 Python 脚本路径
	)
set PythonFile="%~f0"
set PythonFile=%PythonFile:.cmd=.py%
set PythonFile=%PythonFile:.bat=.py%
title %PythonFile:"=%
if %test%==1 echo %PythonFile:"=%


REM 检测脚本存在与否
if not exist %PythonFile% (
	if %test%==1 (
		echo.
		echo 执行脚本
		call "%PythonPath%"
		)^
	else (
		echo 脚本不存在，退出运行
		echo.
		goto End
		)
	)


REM 执行脚本
if %test%==1 (
	echo.
	echo 执行脚本
	)
echo USE %PythonPath:"=%
echo RUN %PythonFile:"=%
echo.
echo.
call "%PythonPath%" %PythonFile%
goto End


:End
if exist %PythonPathTxt% (
	del /q %PythonPathTxt%
	)
pause & exit

