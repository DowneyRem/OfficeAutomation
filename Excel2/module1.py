#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os, sys
from win32com.client import DispatchEx, Dispatch
from functools import wraps


def timethis(func):
	@wraps(func)
	def wrapper(*args, **kwargs):
		start = time.perf_counter()
		r = func(*args, **kwargs)
		end = time.perf_counter()
		print('{}.{} : {}'.format(func.__module__, func.__name__, end - start))
		return r
	return wrapper


def openfileCheck(func):
	@wraps(func)
	def wrapper(*args, **kwargs):
		arg = args[0]
		if os.path.exists(arg):
			try:
				r = func(*args, **kwargs)
				return r
			except IOError as e:
				print("文件被占用：{}".format(arg))
		else:
			print("文件不存在：{}".format(arg))
	return wrapper


@timethis
@openfileCheck
def editExcel(path, Visible=0):  # 打开软件自动操作
	excel = Dispatch('Excel.Application')  # 工作簿共用进程
	excel.Visible = Visible  # 0为后台运行
	excel.DisplayAlerts = 0  # 不显示警告
	print("处理数据中……")
	wb = excel.Workbooks.Open(path, UpdateLinks=3)  # 打开文档；更新链接
	# wb.Application.Run("自动化")  # 运行宏
	
	file = wb.Path + "/" + wb.Name
	if  "sharepoint.com" in wb.Path:
		# 获取OneDrive在线链接，截取掉前段云端路径
		file = file.replace(file[:84], "")
		file = os.environ.get("OneDrive") + file
		
	file = file.replace("xlsm", "xlsx")
	# print(file)
	
	if "唐门小说更新统计.xlsm" in wb.Name:
		# 选中单元格，输入上次保存时间，保存
		lastsave = wb.BuiltinDocumentProperties("Last Save Time")
		lastsave = str(lastsave).replace("-", "/")[:10]
		wb.Worksheets("小说").Range("J1").Value = "更新时间: " + lastsave
		
		# 使用SaveCopyAs 无需设置文件格式
		wb.Save()
		wb.SaveCopyAs(file)
		print("文件已保存：{}".format(wb.Name))
		
	else:
		print ("文件错误！\n当前文件：" + wb.Name)


if __name__ == '__main__':
	pass
