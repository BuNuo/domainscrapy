# -*- coding: utf-8 -*-

# Define your item pipelines here
#
# Don't forget to add your pipeline to the ITEM_PIPELINES setting
# See: https://doc.scrapy.org/en/latest/topics/item-pipeline.html
from xlrd import open_workbook
from xlutils.copy import copy


class DomainscrapyPipeline(object):
    def process_item(self, item, spider):

        # line = [item['domain'], item['registrar'][0], item['email'][0],item['phone'][0],item['create_time'][0],
        #         item['expire_time'][0], item['server'][0], item['status'][0]]
        # fp = open('../tmp/domain_new.txt', 'a', encoding='utf-8')
        # fp.write(','.join(line))
        # fp.write('/n')
        # fp.close()

        if item:
            domain = item['domain']
            registrar = item['registrar'][0] if item['registrar'] else item['registrar']
            email = item['email'][0] if item['email'] else item['email']
            phone = item['phone'][0] if item['phone'] else item['phone']
            create_time = item['create_time'][0] if item['create_time'] else item['create_time']
            expire_time = item['expire_time'][0] if item['expire_time'] else item['expire_time']
            server = item['server'][0] if item['server'] else item['server']
            status = ','.join(item['status']) if item['status'] else item['status']



            # book = xlrd.open_workbook(r'domainInfo.xls')
            # wfile = copy(book)  # 在内存中复制
            # wsheet = wfile.get_sheet(0)
            # wsheet.write(2, 1, 'new_new_new')  # 在第3行第2列添加新数据'new_new_new'
            # wfile.save(r'domainInfo.xls')

            rexcel = open_workbook(r'domainInfo.xls')  # 用wlrd提供的方法读取一个excel文件
            rows = rexcel.sheets()[0].nrows  # 用wlrd提供的方法获得现在已有的行数
            excel = copy(rexcel)  # 用xlutils提供的copy方法将xlrd的对象转化为xlwt的对象
            table = excel.get_sheet(0)  # 用xlwt对象的方法获得要操作的sheet

            table.write(rows, 0, rows)
            table.write(rows, 1, domain)
            table.write(rows, 2, registrar)
            table.write(rows, 3, email)
            table.write(rows, 4, phone)
            table.write(rows, 5, create_time)
            table.write(rows, 6, expire_time)
            table.write(rows, 7, status)
            # print("table: ",table)

            excel.save(r'domainInfo.xls')

        return item
