@echo off
REM echo 切换至 UTF8
chcp 65001 >nul
REM echo 开启延迟变量
setlocal EnableDelayedExpansion
REM test=1 输出测试内容
set test=0


call :myecho 当前目录
echo %~dp0
cd /d %~dp0
set PythonPathTxt=%~dp0PythonPath.txt
echo.
goto FindPython
EXIT /B %ERRORLEVEL% 


:FindPython
call :myecho 寻找Python程序路径
dir /s/b | findstr "python.exe" >> %PythonPathTxt%
for /f "usebackq delims=," %%a in (%PythonPathTxt%) do (
	set "PythonPath=%%~a"
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
	call :myecho 当前目录下没有python.exe 
	goto BackToParentPath
	)^
else (
	if %test%==1 echo.
	call :myecho 已经找到python.exe
	call :myecho "!PythonPath:"=!"
	goto RunScript
	)
EXIT /B %ERRORLEVEL% 


:BackToParentPath
if %test%==1 echo.
call :myecho 返回上级目录
for %%a in ("%cd%") do set ParentPath=%%~dpa
call :myecho %ParentPath%
cd %ParentPath%
goto FindPython
EXIT /B %ERRORLEVEL% 


:RunScript
if %test%==1 echo.
call :myecho 寻找Python脚本路径
set PythonFile="%~f0"
set PythonFile=%PythonFile:.cmd=.py%
set PythonFile=%PythonFile:.bat=.py%
title %PythonFile:"=%
call :myecho %PythonFile:"=%


if %test%==1 echo.
call :myecho 执行脚本
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
EXIT /B %ERRORLEVEL% 


:End
pause & exit
EXIT /B %ERRORLEVEL% 


:myecho 
REM echo myecho
if %test%==1 echo %~1
EXIT /B 0