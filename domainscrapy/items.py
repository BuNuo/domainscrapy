# -*- coding: utf-8 -*-

# Define here the models for your scraped items
#
# See documentation in:
# https://doc.scrapy.org/en/latest/topics/items.html

import scrapy


class DomainscrapyItem(scrapy.Item):
    # define the fields for your item here like:
    # name = scrapy.Field()
    # regist = scrapy.Field()
    # date = scrapy.Field()
    # dns = scrapy.Field()
    # contect = scrapy.Field()
    # domain = scrapy.Field()
    # status = scrapy.Field()
    registrar = scrapy.Field()
    email = scrapy.Field()
    phone = scrapy.Field()
    create_time = scrapy.Field()
    expire_time = scrapy.Field()
    server = scrapy.Field()
    domain = scrapy.Field()
    status = scrapy.Field()
