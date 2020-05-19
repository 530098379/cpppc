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
import json

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

	headers ={
		"content-type":"application/json"
	}
	# 获取当前工会的所有年报
	url_union = "https://www.cpppc.org:8082/api/pub/project/search"
	param_union = {"created_date_order":"desc","dist_city":"","dist_code":"",
			"dist_province":"","end":"","industry":"",
			"level":"","max":"10000000000000000","min":"0",
			"name":"","pageNumber":"1","size":"5","start":"",
			"status":["0","1","2"]}
	r =requests.post(url_union, headers = headers, data = json.dumps(param_union))

	if r.status_code != 200:
		raise Exception(r.status_code)

	json_data = json.loads(r.text)

	# 再次封装，获取具体标签内的内容
	for ppp_data in json_data["data"]["hits"]:
		print("工会编号:" + ppp_data["proj_rid"], flush = True)
		print("--------------------------")
		sheet.write(count,0, ppp_data["proj_rid"]) # row, column, value
		workbook.save(excel_file_name)
		count = count + 1;
		time.sleep(2)

	print("完成",flush = True)
