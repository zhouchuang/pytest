import scrapy

class TorentItem(scrapy.Item):
    url = scrapy.Field()
    name= scrapy.Field()
    description  = scrapy.Field()
    size = scrapy.Field()
    print  url