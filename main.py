#!/usr/bin/env python3

import requests
from pathlib import Path
import io
import sys
import re
import xlwt
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
	# 获取每页的项目id
	proc_url = "https://www.cpppc.org:8082/api/pub/project/search"
	proc_param = {"created_date_order":"desc","dist_city":"","dist_code":"",
			"dist_province":"","end":"","industry":"",
			"level":"","max":"10000000000000000","min":"0",
			"name":"","pageNumber":"1","size":"5","start":"",
			"status":["0","1","2"]}
	proc_r =requests.post(proc_url, headers = headers, data = json.dumps(proc_param))

	if proc_r.status_code != 200:
		raise Exception(proc_r.status_code)

	proc_json_data = json.loads(proc_r.text)

	# 根据项目id，获取对应的数据
	for proc_data in proc_json_data["data"]["hits"]:
		pay_url = "https://www.cpppc.org:8082/api/pub/project/prepare-detail/" + proc_data["proj_rid"]
		pay_r =requests.get(pay_url, headers = headers)
		if pay_r.status_code != 200:
			raise Exception(pay_r.status_code)

		pay_json_data = json.loads(pay_r.text)
		
		for pay_data in pay_json_data["data"]["prepareFinancial"]["payDutyRatioList"]:
			sheet.write(count,0, proc_data["proj_no"]) # row, column, value
			sheet.write(count,1, pay_data["year"])
			sheet.write(count,2, pay_data["ratioA"]/1000000)
			sheet.write(count,3, pay_data["ratioA"]/1000000)
			sheet.write(count,4, pay_data["ratioE"]/1000000)
			sheet.write(count,5, pay_data["ratioG"]/1000000)
			sheet.write(count,6, pay_data["ratio"])
			workbook.save(excel_file_name)
			count = count + 1;
			time.sleep(2)

	print("完成",flush = True)
