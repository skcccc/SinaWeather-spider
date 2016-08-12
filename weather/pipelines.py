# -*- coding: utf-8 -*-
# Define your item pipelines here
# Don't forget to add your pipeline to the ITEM_PIPELINES setting
# See: http://doc.scrapy.org/en/latest/topics/item-pipeline.html

import xlrd
import xlwt


class WeatherPipeline(object):
	def __init__(self):
		pass

	def process_item(self, item, spider):
		file = open('wea.txt', 'w+')
		workbook = xlwt.Workbook()
		sheet = workbook.add_sheet("Sheet1")
		city = item['city']
		file.write('city:' + str(city) + '\n\n')

		date = item['date']

		desc = item['dayDesc']
		dayDesc = desc[1::2]
		nightDesc = desc[0::2]

		dayTemp = item['dayTemp']

		weaitem_t = zip(date, dayDesc, nightDesc, dayTemp)
		weaitem = list(weaitem_t)

		for i in range(len(weaitem)):
			item = weaitem[i]
			d = item[0]
			dd = item[1]
			nd = item[2]
			ta = item[3].split('/')
			dt = ta[0]
			nt = ta[1]
			witem = [d,dd,dt,nd,nt]
			titem = ['date','daydes', 'daytemp', 'nightdes', 'nighttemp']
			txt = 'date:{0}\t\tday:{1}({2})\t\tnight:{3}({4})\n\n'.format(d, dd, dt, nd, nt)

			for n in range(5):
				if i == 0:
					sheet.write(i, n, titem[n])
				else:
					sheet.write(i, n, witem[n])
			file.write(txt)
			workbook.save(str(city)+'wea.xls')

		return item
