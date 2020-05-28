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
	pay_count = 1
	pay_workbook = xlwt.Workbook()
	pay_sheet = pay_workbook.add_sheet("Sheet Name1")
	pay_sheet.write(0, 0, "编号") # row, column, value
	pay_sheet.write(0, 1, "支出年度")
	pay_sheet.write(0, 2, "A-本项目一般公共预算支出数额（万元）")
	pay_sheet.write(0, 3, "B-本级已有管理库项目一般公共预算支出数额（万元）")
	pay_sheet.write(0, 4, "C-年度一般公共预算 支出数额（万元）")
	pay_sheet.write(0, 5, "占比（%）")

	base_count = 1
	base_workbook = xlwt.Workbook()
	base_sheet = base_workbook.add_sheet("Sheet Name1")
	base_sheet.write(0, 0, "项目名称") # row, column, value
	base_sheet.write(0, 1, "编号")
	base_sheet.write(0, 2, "所在区域")
	base_sheet.write(0, 3, "所属行业")
	base_sheet.write(0, 4, "项目总投资")
	base_sheet.write(0, 5, "所处阶段")
	base_sheet.write(0, 6, "发起时间")
	base_sheet.write(0, 7, "项目示范级别/批次")
	base_sheet.write(0, 8, "回报机制")
	base_sheet.write(0, 9, "项目联系人")
	base_sheet.write(0, 10, "联系电话")
	base_sheet.write(0, 11, "合作期限")
	base_sheet.write(0, 12, "运作方式")
	base_sheet.write(0, 13, "采购方式")
	base_sheet.write(0, 14, "指标数量")
	for i in range(1, 15):
		base_sheet.write(0, i * 3 + 12, "指标" + str(i))
		base_sheet.write(0, i * 3 + 13, "权重" + str(i))
		base_sheet.write(0, i * 3 + 14, "评分" + str(i))

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
		print("当前页数:" + str(pageNumber) + "/" + str(proc_count), flush = True)
		print("--------------------------")
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

			# 采购方式
			purchase_way = pay_json_data["data"]["implPlanInfo"]["socialPurchaseWay"]

			# 预算指标数据
			for pay_data in pay_json_data["data"]["prepareFinancial"]["payDutyRatioList"]:
				pay_sheet.write(pay_count, 0, proc_data["proj_no"]) # row, column, value
				pay_sheet.write(pay_count, 1, pay_data["year"] if "year" in pay_data else "")
				pay_sheet.write(pay_count, 2, pay_data["ratioA"]/1000000 if "ratioA" in pay_data else "")
				pay_sheet.write(pay_count, 3, pay_data["ratioE"]/1000000 if "ratioE" in pay_data else "")
				pay_sheet.write(pay_count, 4, pay_data["ratioG"]/1000000 if "ratioG" in pay_data else "")
				pay_sheet.write(pay_count, 5, pay_data["ratio"] if "ratio" in pay_data else "")
				pay_workbook.save(pay_excel_file_name)
				pay_count = pay_count + 1;

			base_url = "https://www.cpppc.org:8082/api/pub/project/detail/" + proc_data["proj_rid"]
			base_r = requests.get(base_url, headers = headers)
			if base_r.status_code != 200:
				raise Exception(base_r.status_code)

			base_json_data = json.loads(base_r.text)
			base_data = base_json_data["data"]

			# 基本指标数据
			# 项目名称
			base_sheet.write(base_count, 0, base_data["projName"])

			# 编号
			base_sheet.write(base_count, 1, base_data["projNo"])

			# 所在区域
			base_sheet.write(base_count, 2, base_data["distProvinceName"] \
				+ " - " + base_data["distCityName"] \
				+ (" - " + base_data["distName"] if base_data["distName"] else "") )

			# 所属行业
			base_sheet.write(base_count, 3, base_data["industryRequiredName"] \
				+ " - " + base_data["industryOptionalName"])

			# 项目总投资
			base_sheet.write(base_count, 4, str(base_data["investCount"]/1000000) + "万元")

			# 所处阶段
			if base_data["projState"] == "1":
				base_sheet.write(base_count, 5, "准备阶段")
			elif base_data["projState"] == "2":
				base_sheet.write(base_count, 5, "采购阶段")
			elif base_data["projState"] == "3":
				base_sheet.write(base_count, 5, "执行阶段")
			else:
				base_sheet.write(base_count, 5, "")

			# 发起时间
			base_sheet.write(base_count, 6, base_data["startTime"])

			# 项目示范级别/批次
			if base_data["example"] == 1:
				batch_level = ""
				for example_data in base_data["exampleList"]:
					batch_name = ""
					batch_number = example_data["batchNumber"]
					if batch_number == "1":
						batch_name = "第一批次"
					elif batch_number == "2":
						batch_name = "第二批次"
					elif batch_number == "3":
						batch_name = "第三批次"
					elif batch_number == "4":
						batch_name = "第四批次"
					elif batch_number == "-1":
						batch_name = "无批次信息"
					else:
						batch_name = ""
					
					example_level = example_data["exampleLevel"]
					example_level_name = ""
					if example_level == "0":
						example_level_name = "国家级"
					elif example_level == "1":
						example_level_name = "省级"
					elif example_level == "2":
						example_level_name = "市级"
					else:
						example_level_name = ""
				batch_level += batch_name
				batch_level += example_level_name

				base_sheet.write(base_count, 7, batch_level)
			elif base_data["example"] == 2:
				base_sheet.write(base_count, 7, "暂无")
			else:
				base_sheet.write(base_count, 7, "")

			# 回报机制
			if base_data["returnMode"] == "1":
				base_sheet.write(base_count, 8, "政府付费")
			elif base_data["returnMode"] == "2":
				base_sheet.write(base_count, 8, "使用者付费")
			elif base_data["returnMode"] == "3":
				base_sheet.write(base_count, 8, "可行性缺口补助")
			else:
				base_sheet.write(base_count, 8, "")

			# 项目联系人
			base_sheet.write(base_count, 9, base_data["linkUname"])

			# 联系电话
			base_sheet.write(base_count, 10, base_data["linkTel"])

			# 合作期限
			base_sheet.write(base_count, 11, str(base_data["cooperationTerm"]) + "年")

			# 运作方式
			if base_data["operateMode"] == "1":
				base_sheet.write(base_count, 12, "BOT")
			elif base_data["operateMode"] == "2":
				base_sheet.write(base_count, 12, "TOT")
			elif base_data["operateMode"] == "3":
				base_sheet.write(base_count, 12, "ROT")
			elif base_data["operateMode"] == "4":
				base_sheet.write(base_count, 12, "BOO")
			elif base_data["operateMode"] == "5":
				base_sheet.write(base_count, 12, "TOT+BOT")
			elif base_data["operateMode"] == "6":
				base_sheet.write(base_count, 12, "TOT+BOO")
			elif base_data["operateMode"] == "7":
				base_sheet.write(base_count, 12, "OM")
			elif base_data["operateMode"] == "8":
				base_sheet.write(base_count, 12, "MC")
			elif base_data["operateMode"] == "9":
				base_sheet.write(base_count, 12, "其他")
			else:
				base_sheet.write(base_count, 12, "")

			# 采购方式
			if purchase_way == "1":
				base_sheet.write(base_count, 13, "公开招标")
			elif purchase_way == "2":
				base_sheet.write(base_count, 13, "竞争性谈判")
			elif purchase_way == "3":
				base_sheet.write(base_count, 13, "邀请招标")
			elif purchase_way == "4":
				base_sheet.write(base_count, 13, "竞争性磋商")
			elif purchase_way == "5":
				base_sheet.write(base_count, 13, "单一来源采购")
			else:
				base_sheet.write(base_count, 13, "")

			# 评价指标的数量
			base_sheet.write(base_count, 14, len(pay_json_data["data"]["prepareValue"]["projectPreValueEvaList"]))

			# 权重数据
			base_for_count = 15

			for quanzhong_data in pay_json_data["data"]["prepareValue"]["projectPreValueEvaList"]:
				base_sheet.write(base_count, base_for_count, \
					quanzhong_data["indicatorName"] if "indicatorName" in quanzhong_data else "")

				base_sheet.write(base_count, base_for_count + 1, \
					quanzhong_data["weight"] if "weight" in quanzhong_data else "")

				base_sheet.write(base_count, base_for_count + 2, \
					quanzhong_data["scoreResult"] if "scoreResult" in quanzhong_data else "")

				base_for_count = base_for_count + 3

			base_workbook.save(base_excel_file_name)

			base_count = base_count + 1;
			time.sleep(2)

	print("完成", flush = True)
	print("--------------------------")
