# -*- coding: utf-8 -*-
import scrapy
import xlrd
from domainscrapy.items import DomainscrapyItem


class ChinazSpider(scrapy.Spider):
    name = 'chinaz'
    allowed_domains = ['chinaz.com']
    start_urls = []

    workbook = xlrd.open_workbook(r'domains.xlsx')
    sheet = workbook.sheet_by_index(0)

    # print('总数:', sheet.nrows - 1)

    suffix = ['com', 'cn', 'org', 'net', 'cc']

    for i in range(sheet.nrows):
        if i == 0:
            continue
        domainValue = sheet.row_values(i)[0]
        for j in suffix:
            domain = domainValue + '.' + j
            url = 'http://whois.chinaz.com/' + domain
            start_urls.append(url)

    def clean_str(self):
        data = self.data
        data = [t.strip() for t in data]
        data = [t.strip('\r') for t in data]
        data = [t for t in data if t != '']
        #data = [t.encode('utf-8') for t in data]
        return data

    def parse(self, response):
        surl = response.url
        domain = surl.split('/')[3].strip()

        item = DomainscrapyItem()

        if response.xpath('//ul[@class="WhoisLeft fl"]/li'):
            domainRemarks = []
            for r in response.xpath('//ul[@class="WhoisLeft fl"]/li'):
                domainRemark = r.xpath('./div[1]/text()').extract()
                # domainRemarks.append(domainRemark[0]) if len(domainRemark) else domainRemarks
                domainRemarks.append(domainRemark[0]) if len(domainRemark) else domainRemarks.append('')

            # print(domainRemarks)
            # domain = response.xpath('//ul[@class="WhoisLeft fl"]/li[1]/div[2]/p[1]/a[1]/text()').extract()
            # print("域名:", domain)

            try:
                index = domainRemarks.index('注册商')
                registrar = response.xpath(
                    '//ul[@class="WhoisLeft fl"]/li[' + str(index + 1) + ']/div[2]/div/span/text()').extract()
                # print("注册商:", registrar[0])
            except:
                registrar = []
            # contact = response.xpath('//ul[@class="WhoisLeft fl"]/li[3]/div[2]/span/text()').extract()
            # print("联系人:",contact[0])

            try:
                index = domainRemarks.index('联系邮箱')
                email = response.xpath(
                    '//ul[@class="WhoisLeft fl"]/li[' + str(index + 1) + ']/div[2]/span/text()').extract()
                # print("邮箱:", email[0])
            except:
                email = []

            try:
                index = domainRemarks.index('联系电话')
                phone = response.xpath(
                    '//ul[@class="WhoisLeft fl"]/li[' + str(index + 1) + ']/div[2]/span/text()').extract()
                # print("电话:", phone[0])
            except:
                phone = []

            # update_time = response.xpath('//ul[@class="WhoisLeft fl"]/li[6]/div[2]/span/text()').extract()
            # print("更新时间:", update_time[0])

            try:
                index = domainRemarks.index('创建时间')
                create_time = response.xpath(
                    '//ul[@class="WhoisLeft fl"]/li[' + str(index + 1) + ']/div[2]/span/text()').extract()
                # print("创建时间:", create_time[0])
            except:
                create_time = []

            try:

                index = domainRemarks.index('过期时间')
                expire_time = response.xpath(
                    '//ul[@class="WhoisLeft fl"]/li[' + str(index + 1) + ']/div[2]/span/text()').extract()
                # print("过期时间:", expire_time[0])
            except:
                expire_time = []

            # company = response.xpath('//ul[@class="WhoisLeft fl"]/li[8]/div[2]/span/text()').extract()
            # print("公司:", company)

            try:
                index = domainRemarks.index('域名服务器')
                server = response.xpath(
                    '//ul[@class="WhoisLeft fl"]/li[' + str(index + 1) + ']/div[2]/span/text()').extract()
                # print("域名服务器:", server[0])
            except:
                server = []

            # dns = response.xpath('//ul[@class="WhoisLeft fl"]/li[10]/div[2]/text()').extract()
            # print("DNS:", dns)

            try:
                index = domainRemarks.index('状态')
                status = []
                for s in response.xpath('//ul[@class="WhoisLeft fl"]/li[' + str(index + 1) + ']/div[2]/p'):
                    state = s.xpath('./span/a/text()').extract()
                    status.append(state[0])
                    # print("状态:", status)
            except:
                status = []

            item['registrar'] = registrar
            item['email'] = email
            item['phone'] = phone
            item['create_time'] = create_time
            item['expire_time'] = expire_time
            item['server'] = server
            item['domain'] = domain
            item['status'] = status
            # print(item)

        else:
            print('可用')

        return item
