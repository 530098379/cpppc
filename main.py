#!/usr/bin/env python3

import requests
from pathlib import Path
import io
import sys
import re
import xlwt
import xlrd
import os
import time
import datetime

if __name__ == "__main__":
	sys.stdout = io.TextIOWrapper(sys.stdout.buffer,encoding='utf-8')
	print("开始", flush = True)
	# 文件名
	excel_file_name = os.getcwd() + "\\result_" + \
		datetime.datetime.now().strftime("%Y%m%d%H%M%S") + ".xls"

	# 做成Excel文件
	count=0
	workbook = xlwt.Workbook()
	sheet = workbook.add_sheet("Sheet Name1")

	# 获取cookie
	url_cok = "https://www.cpppc.org:8082/inforpublic/homepage.html#/searchresult"
	r_cok =requests.get(url_cok)
	cookie_jar = r_cok.cookies

	if r_cok.status_code != 200:
		raise Exception(r_cok.status_code)

	# 获取当前工会的所有年报
	url_union = "https://www.cpppc.org:8082/api/pub/project/search"
	param_union = {"created_date_order":"desc","dist_city":"","dist_code":"",
			"dist_province":"","end":"","industry":"",
			"level":"","max":"10000000000000000","min":"0",
			"name":"","pageNumber":"1","size":"5","start":"",
			"status":["0","1","2"]}
	r =requests.post(url_union, param_union, cookies=cookie_jar)

	if r.status_code != 200:
		raise Exception(r.status_code)

	# 再次封装，获取具体标签内的内容
	for ppp_data in r.json():
		print("工会编号:" + ppp_data["aaa"], flush = True)
		print("年份:" + ppp_data["aaa"], flush = True)
		print("内容:" + ppp_data["aaa"], flush = True)
		print("--------------------------")
		sheet.write(count,0, ppp_data["aaa"]) # row, column, value
		sheet.write(count,1, ppp_data["aaaa"])
		sheet.write(count,2, ppp_data["aaaa"])

	print("完成",flush = True)
