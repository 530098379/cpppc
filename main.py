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
	print("--------------------------")

	deta_str = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
	# 文件名
	pay_excel_file_name = os.getcwd() + "\\预算指标_" + deta_str + ".xls"

	base_excel_file_name = os.getcwd() + "\\基本指标_" + deta_str + ".xls"

	# 做成Excel文件
	pay_count=0
	pay_workbook = xlwt.Workbook()
	pay_sheet = pay_workbook.add_sheet("Sheet Name1")

	base_count=0
	base_workbook = xlwt.Workbook()
	base_sheet = base_workbook.add_sheet("Sheet Name1")

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

		print("读取数据")
		print("项目编号:" + proc_data["proj_no"], flush = True)
		print("项目名称:" + proc_data["proj_name"], flush = True)
		print("--------------------------")

		pay_json_data = json.loads(pay_r.text)
		
		# 预算指标数据
		for pay_data in pay_json_data["data"]["prepareFinancial"]["payDutyRatioList"]:
			pay_sheet.write(pay_count,0, proc_data["proj_no"]) # row, column, value
			pay_sheet.write(pay_count,1, pay_data["year"])
			pay_sheet.write(pay_count,2, pay_data["ratioA"]/1000000)
			pay_sheet.write(pay_count,3, pay_data["ratioA"]/1000000)
			pay_sheet.write(pay_count,4, pay_data["ratioE"]/1000000)
			pay_sheet.write(pay_count,5, pay_data["ratioG"]/1000000)
			pay_sheet.write(pay_count,6, pay_data["ratio"])
			pay_workbook.save(pay_excel_file_name)
			pay_count = pay_count + 1;
		
		# 基本指标数据
 		# 编号
		base_sheet.write(base_count,0, proc_data["proj_no"])

		# 所在区域
		base_sheet.write(base_count,1, proc_data["dist_province_name"] \
			+ " - " + proc_data["dist_city_name"] \
			+ (" - " + proc_data["dist_code_name"] if proc_data["dist_code_name"] else "") )

		# 所属行业
		base_sheet.write(base_count,2, proc_data["industry_required_name"] \
			+ " - " + proc_data["industry_optional_name"])

		# 项目总投资
		base_sheet.write(base_count,3, proc_data["invest_count"]/1000000)

		# 所处阶段
		base_sheet.write(base_count,4, "")

		# 发起时间
		base_sheet.write(base_count,5, "")

		# 项目示范级别/批次
		base_sheet.write(base_count,6, "")

		# 回报机制
		base_sheet.write(base_count,7, "")

		# 项目联系人
		base_sheet.write(base_count,8, "")

		# 联系电话
		base_sheet.write(base_count,9, "")

		# 合作期限
		base_sheet.write(base_count,10, "")

		# 运作方式
		base_sheet.write(base_count,11, "")

		# 采购方式
		base_sheet.write(base_count,12, "")

		base_for_count = 13
		for base_data in pay_json_data["data"]["prepareValue"]["projectPreValueEvaList"]:
			base_sheet.write(base_count,base_for_count, base_data["indicatorName"])# row, column, value
			base_sheet.write(base_count,base_for_count + 1, base_data["weight"])
			base_sheet.write(base_count,base_for_count + 2, base_data["scoreResult"])
			base_for_count = base_for_count + 3

		base_workbook.save(base_excel_file_name)

		base_count = base_count + 1;
		time.sleep(5)


	print("完成",flush = True)
	print("--------------------------")
