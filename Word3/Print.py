#!/usr/bin/python
# -*- coding: UTF-8 -*-
import os
import re
import time
import shutil
import winreg
from functools import wraps
from platform import platform

from win32com.client import Dispatch, DispatchEx


def timethis(func):
	@wraps(func)
	def wrapper(*args, **kwargs):
		start = time.perf_counter()
		r = func(*args, **kwargs)
		end = time.perf_counter()
		print('{}.{} : {}'.format(func.__module__, func.__name__, end - start))
		return r
	return wrapper


def openFileCheck(func):
	@wraps(func)
	def wrapper(*args, **kwargs):
		arg = args[0]
		if os.path.exists(arg):
			try:
				r = func(*args, **kwargs)
				return r
			except IOError:
				print("文件被占用：{}".format(arg))
		else:
			print("文件不存在：{}".format(arg))
	return wrapper


def saveFileCheck(func):
	@wraps(func)
	def wrapper(*args, **kwargs):
		arg = args[0]
		(dir, name) = os.path.split(arg)  # 分离文件名和目录名
		if not os.path.exists(dir):
			os.makedirs(dir)
		name = formatFileName(name)
		path = os.path.join(dir, name)
		r = func(*args, **kwargs)
		return r
	return wrapper


# 已通过@saveFileCheck加入保存文件的函数内
def formatFileName(text):
	if text:
		text = re.sub('[/\:*?"<>|]', ' ', text)
		text = text.replace('  ', '')
		return text


pathlist = []
def findFile(path, *extnames):
	# 省略 extnames 参数可以获取全部文件
	# extname="" 获取无后缀名文件
	for dir in os.listdir(path):
		dir = os.path.join(path, dir)
		if os.path.isdir(dir):
			findFile(dir, *extnames)
		
		if os.path.isfile(dir):
			if len(extnames) > 0:
				for extname in extnames:
					(name, ext) = os.path.splitext(dir)
					if ext == extname:
						pathlist.append(dir)
			elif len(extnames) == 0:
				pathlist.append(dir)
	return pathlist


@timethis
@openFileCheck
def openWord(path):
	word = Dispatch('Word.Application')  # 非独立进程
	word.Visible = 1  # 0为后台运行
	word.DisplayAlerts = 0  # 不显示，不警告
	docx = word.Documents.Open(path)   # 打开文档
	# print("打开Word……")
	docx.Application.Run("打印")  # 运行宏
	# print("打印中……")
	docx.Close(True)
	# time.sleep(1)
	# word.Quit(1)
	# word.Quit()


@openFileCheck
def openExcel(path):  # 打开软件手动操作
	excel = DispatchEx('Excel.Application')  # 独立进程
	excel.Visible = 1  # 0为后台运行
	excel.DisplayAlerts = 0  # 不显示，不警告
	xlsx = excel.Workbooks.Open(path)  # 打开文档
	print("打开Excel……")


def desktop():
	if "Windows" in platform():  # 其他平台我也没用过
		key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
							 r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')
		return winreg.QueryValueEx(key, "Desktop")[0]
		
	
@timethis	
def main():
	path1, path2 = "", ""
	path = os.path.join(desktop(),"path1","path2")
	pathlist = findFile(path, ".doc", ".docx")
	pathlist.reverse()
	for i in range(len(pathlist)):
		file = pathlist[i]
		name = os.path.basename(file)
		print(f"{len(pathlist)-i}, {name}")

		if len(pathlist)-i > 1000:
			continue
		elif len(pathlist)-i > 0:
			openWord(file)
			time.sleep(5)
		else:
			break
			
	
	
if __name__ == "__main__":
	main()
	pass
	
	