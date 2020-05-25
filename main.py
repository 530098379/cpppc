#!/usr/bin/env python3

import requests
import io
import sys
import xlwt
import os
import time
import datetime
import json
import math

if __name__ == "__main__":
	sys.stdout = io.TextIOWrapper(sys.stdout.buffer,encoding='utf-8')
	print("开始", flush = True)
	print("--------------------------")

	deta_str = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
	# 文件名
	pay_excel_file_name = os.getcwd() + "\\预算指标_" + deta_str + ".xls"

	base_excel_file_name = os.getcwd() + "\\基本指标_" + deta_str + ".xls"

	# 做成Excel文件
	pay_count = 0
	pay_workbook = xlwt.Workbook()
	pay_sheet = pay_workbook.add_sheet("Sheet Name1")

	base_count = 0
	base_workbook = xlwt.Workbook()
	base_sheet = base_workbook.add_sheet("Sheet Name1")

	headers = {
		"content-type":"application/json"
	}

	# 获取项目总数
	url = "https://www.cpppc.org:8082/api/report/open/mapdemooverall?divisionCode=0"
	r =requests.get(url, headers = headers)
	if r.status_code != 200:
		raise Exception(r.status_code)

	count_data = json.loads(r.text)
	proc_count = math.ceil(int(count_data["data"]["project"]["totalCnt"])/5)
	print("总页数:" + str(proc_count), flush = True)
	print("--------------------------")

	for pageNumber in range(1, proc_count + 1):
		# 获取每页的项目id
		proc_url = "https://www.cpppc.org:8082/api/pub/project/search"
		proc_param = {"created_date_order":"desc","dist_city":"","dist_code":"",
				"dist_province":"","end":"","industry":"",
				"level":"","max":"10000000000000000","min":"0",
				"name":"","pageNumber":pageNumber,"size":"5","start":"",
				"status":["0","1","2"]}
		proc_r = requests.post(proc_url, headers = headers, data = json.dumps(proc_param))

		if proc_r.status_code != 200:
			raise Exception(proc_r.status_code)

		proc_json_data = json.loads(proc_r.text)

		# 根据项目id，获取对应的数据
		for proc_data in proc_json_data["data"]["hits"]:
			pay_url = "https://www.cpppc.org:8082/api/pub/project/prepare-detail/" + proc_data["proj_rid"]
			pay_r = requests.get(pay_url, headers = headers)
			if pay_r.status_code != 200:
				raise Exception(pay_r.status_code)

			print("读取数据")
			print("项目编号:" + proc_data["proj_no"], flush = True)
			print("项目名称:" + proc_data["proj_name"], flush = True)
			print("--------------------------")

			pay_json_data = json.loads(pay_r.text)
			
			# 预算指标数据
			for pay_data in pay_json_data["data"]["prepareFinancial"]["payDutyRatioList"]:
				pay_sheet.write(pay_count, 0, proc_data["proj_no"]) # row, column, value
				pay_sheet.write(pay_count, 1, pay_data["year"])
				pay_sheet.write(pay_count, 2, pay_data["ratioA"]/1000000)
				pay_sheet.write(pay_count, 3, pay_data["ratioA"]/1000000)
				pay_sheet.write(pay_count, 4, pay_data["ratioE"]/1000000)
				pay_sheet.write(pay_count, 5, pay_data["ratioG"]/1000000)
				pay_sheet.write(pay_count, 6, pay_data["ratio"])
				pay_workbook.save(pay_excel_file_name)
				pay_count = pay_count + 1;

			base_url = "https://www.cpppc.org:8082/api/pub/project/detail/" + proc_data["proj_rid"]
			base_r = requests.get(base_url, headers = headers)
			if base_r.status_code != 200:
				raise Exception(base_r.status_code)

			base_json_data = json.loads(base_r.text)
			base_data = base_json_data["data"]

			# 基本指标数据
			# 编号
			base_sheet.write(base_count, 0, base_data["projNo"])

			# 所在区域
			base_sheet.write(base_count, 1, base_data["distProvinceName"] \
				+ " - " + base_data["distCityName"] \
				+ (" - " + base_data["distName"] if base_data["distName"] else "") )

			# 所属行业
			base_sheet.write(base_count, 2, base_data["industryRequiredName"] \
				+ " - " + base_data["industryOptionalName"])

			# 项目总投资
			base_sheet.write(base_count, 3, base_data["investCount"]/1000000)

			# 所处阶段
			if base_data["projState"] == "1":
				base_sheet.write(base_count, 4, "")
			elif base_data["projState"] == "2":
				base_sheet.write(base_count, 4, "")
			else:
				base_sheet.write(base_count, 4, "")

			# 发起时间
			base_sheet.write(base_count, 5, base_data["startTime"])

			# 项目示范级别/批次
			if base_data["projLevel"] == "1":
				base_sheet.write(base_count, 6, "")
			elif base_data["projLevel"] == "3":
				base_sheet.write(base_count, 6, "")
			else:
				base_sheet.write(base_count, 6, "")

			# 回报机制
			if base_data["returnMode"] == "1":
				base_sheet.write(base_count, 7, "")
			elif base_data["returnMode"] == "3":
				base_sheet.write(base_count, 7, "")
			else:
				base_sheet.write(base_count, 7, "")

			# 项目联系人
			base_sheet.write(base_count, 8, base_data["linkUname"])

			# 联系电话
			base_sheet.write(base_count, 9, base_data["linkTel"])

			# 合作期限
			base_sheet.write(base_count, 10, base_data["cooperationTerm"])

			# 运作方式
			if base_data["startType"] == "1":
				base_sheet.write(base_count, 11, "")
			elif base_data["startType"] == "3":
				base_sheet.write(base_count, 11, "")
			else:
				base_sheet.write(base_count, 11, "")

			# 采购方式
			if base_data["operateMode"] == "1":
				base_sheet.write(base_count, 12, "")
			elif base_data["operateMode"] == "3":
				base_sheet.write(base_count, 12, "")
			else:
				base_sheet.write(base_count, 12, "")

			# 评价指标的数量
			base_sheet.write(base_count, 13, base_data["startType"])

			# 权重数据
			base_for_count = 14
			for quanzhong_data in pay_json_data["data"]["prepareValue"]["projectPreValueEvaList"]:
				base_sheet.write(base_count, base_for_count, \
					quanzhong_data["indicatorName"] if quanzhong_data["indicatorName"] else "")
				base_sheet.write(base_count, base_for_count + 1, quanzhong_data["weight"])
				base_sheet.write(base_count, base_for_count + 2, quanzhong_data["scoreResult"])
				base_for_count = base_for_count + 3

			base_workbook.save(base_excel_file_name)

			base_count = base_count + 1;
			time.sleep(3)

	print("完成", flush = True)
	print("--------------------------")
