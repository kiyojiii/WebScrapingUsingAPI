import json
import scrapy   #Scrapy API
import pandas as pd

class SunglassesSpider(scrapy.Spider):
    name = "sunglasses"
    allowed_domains = ["sunglasshut.com"]
    start_urls = [
        "https://www.sunglasshut.com/wcs/resources/plp/10152/byCategoryId/3074457345626651837"
        "?isProductNeeded=true&isChanelCategory=false&pageSize=18&responseFormat=json&currency=USD"
        "&catalogId=20602&top=Y&beginIndex=0&viewTaskName=CategoryDisplayView&storeId=10152&langId=-1"
        "&categoryId=3074457345626651837&pageView=image&orderBy=default&currentPage=1"
    ]

    scraped_data = []

    def parse(self, response):
        data = json.loads(response.body)
        products = data['plpView']['products']['products']['product']

        # Extract and store important data from each product
        for product in products:
            lens_color = product.get('lensColor', 'N/A')
            frame_color = product.get('frameColor', 'N/A')
            seo_currency = product.get('seoCurrency', 'N/A')
            list_price = product.get('listPrice', 'N/A')
            is_out_of_stock = product.get('isOutOfStock', 'N/A')
            lifecycle = product.get('lifecycle', 'N/A')
            is_polarized = product.get('isPolarized', 'N/A')
            part_number = product.get('partNumber', 'N/A')
            unique_id = product.get('uniqueID', 'HIDDEN')

            item = {
                'Product Name': product.get('name', 'N/A'),
                'Price': list_price,
                'Brand': product.get('brand', 'N/A'),
                'Lens Color': lens_color,
                'Frame Color': frame_color,
                'SEO Currency': seo_currency,
                'URL': response.urljoin(product.get('pdpURL', 'N/A')),
                'isOutOfStock': is_out_of_stock,
                'Lifecycle': lifecycle,
                'isPolarized': is_polarized,
                'Part Number': part_number,
                'Unique ID': unique_id
            }
            self.scraped_data.append(item)

        next_page = data['plpView'].get('nextPageURL')
        if next_page is not None:
            next_page = response.urljoin(next_page)
            yield scrapy.Request(next_page, callback=self.parse)
        else:
            # Save the scraped data to an Excel file
            self.save_to_excel()

    def save_to_excel(self):
        # Create a DataFrame from the scraped data
        df = pd.DataFrame(self.scraped_data)

        # Save the DataFrame to an Excel file
        excel_path = 'C:\\Users\\user\\PycharmProjects\\scraped_data\\sunglasses_dataForVideo.xlsx'
        df.to_excel(excel_path, index=False)

# SETUP
# pip install scrapy
# scrapy startproject [name]
# cd [name]
# To run:
# Scrapy crawl sunglasses