#!/usr/bin/python
# -*- coding: UTF-8 -*-
import os
import re
import time
import logging
import webbrowser
from functools import wraps


logging.basicConfig(
	level=logging.INFO,
	format='%(levelname)s %(asctime)s [%(filename)s:%(lineno)d] %(message)s',
	datefmt='%Y.%m.%d. %H:%M:%S',
	# filename='parser_result.log',
	# filemode='w'
)


def timer(function):
	@wraps(function)
	def wrapper(*args, **kwargs):
		start = time.perf_counter()
		result = function(*args, **kwargs)
		end = time.perf_counter()
		print(f"{function.__module__}.{function.__name__}: {end - start}")
		return result
	return wrapper


def openFileCheck(function):
	@wraps(function)
	def wrapper(*args, **kwargs):
		arg = args[0]
		if os.path.exists(arg):
			try:
				result = function(*args, **kwargs)
				return result
			except IOError:
				print(f"文件被占用：{arg}")
		else:
			print(f"文件不存在：{arg}")
	return wrapper


def saveFileCheck(function):
	@wraps(function)
	def wrapper(*args, **kwargs):
		path = args[0]
		(dir, name) = os.path.split(path)
		if not os.path.exists(dir):
			os.makedirs(dir)
		name = formatFileName(name)
		path = os.path.join(dir, name)
		result = function(path, *args[1:], **kwargs)
		return result
	return wrapper


# 已通过@saveFileCheck加入保存文件的函数内
def formatFileName(text) -> str:
	if text:
		text = re.sub('[/:*?"<>|]', ' ', text)
		text = text.replace('  ', '')
		return text
	else:
		return ""


def openFile(path):
	if os.path.exists(path):
		webbrowser.open(path)


@openFileCheck
def openText(path) -> str:
	text = ""
	try:
		with open(path, "r", encoding="UTF8") as f:
			text = f.read()
	except UnicodeError:
		try:
			with open(path, "r", encoding="GB18030") as f:
				text = f.read()
		except UnicodeError:  # Big5 似乎有奇怪的bug，不过目前似乎遇不到
			try:
				with open(path, "r", encoding="BIG5") as f:
					text = f.read()
			except UnicodeError:
				with open(path, "r", encoding="Latin1") as f:
					text = f.read()
	finally:
		return text


@saveFileCheck
def saveText(path, text):
	try:
		with open(path, "w", encoding="UTF8") as f:
			f.write(text)
	except IOError:
		logging.error(f"保存失败：{path}")
	else:
		logging.debug(f"已保存为：{path}")


def replace(text):
	string = "首次进入循环的条件"
	while string:  # 确认输入内容符合正确格式
		string = input("请输入 x0 的值，大于 x0 的坐标均会被替换，按 Enter 键确认：\n")
		# string = input("请输入x0,y0,的值，以英文逗号','为分割符号，按 Enter 键确认：\n")
		try:
			x0 = int(string)
		# x0, y0, *_ = string.split(",")
		# print(x1, y1, _, sep=", ")
		# x0, y0, = int(x0), int(y0)
		except ValueError:
			print("输入有误，请重新输入，退出请按 Enter 键\n")
		else:
			break
	
	if string:  # 开始替换
		count = 0  # 计数器，替换次数
		num1 = 1  # x2 增加量
		num2 = 0  # y2 增加量
		result = re.findall('"position": "([0-9]+),([0-9]+)"', text)
		# print(result)
		for x1, y1 in result:
			count = count + 1
			x2 = int(x1) + num1
			y2 = int(y1) + num2
			if int(x1) > x0:
				# if int(x1) > x0 and int(y1) > y0:
				text = re.sub(f'"position": "{x1},{y1}"', f'"position": "{x2},{y2}"', text)
				print(f"第{count}次替换：已将({x1},{y1})替换为({x2},{y2})")
			else:
				print(f"第{count}次替换：因 {x1} <= {x0}，跳过({x1},{y1})")
	return text


def main(path):
	text = openText(path)
	text = replace(text)
	newpath = os.path.join(os.getcwd(), "Replaced", os.path.basename(path))
	saveText(newpath, text)    # 替换后文件保存路径：程序所在文件夹下的 Replaced 文件夹内
	print(f"\n所有替换均以保存，正在退出，文件路径如下：\n{newpath}")
	time.sleep(1)
	openFile(newpath)


if __name__ == '__main__':
	path = os.path.join(os.getcwd(), "ResearchGraph.json")
	# path = r"D:\Download\ResearchGraph.json"  # 也可以使用被替换文件的绝对路径
	if os.path.exists(path):
		main(path)
	else:
		print("请把程序放在 ResearchGraph.json 所在文件夹内")
		time.sleep(5)
	