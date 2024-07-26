from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import pandas as pd

class AmazonScraper:
    def __init__(self, url):
        self.url = url
        self.driver = self._initialize_driver()
        self.data = []

    def _initialize_driver(self):
        chrome_options = Options()
        chrome_options.add_experimental_option("detach", True)
        driver = webdriver.Chrome(options=chrome_options)
        return driver

    def open_page(self):
        self.driver.get(self.url)

    def scrape_product_details(self):
        cards = self.driver.find_elements("class name", "sg-col-4-of-24")
        for card in cards:
            in_data = []
            is_data = False
            try:
                title = card.find_element("class name", "s-line-clamp-4")
                in_data.append(title.text)
                is_data = True
            except:
                is_data = False

            if is_data:
                try:
                    offer_price = card.find_element("class name", "a-price-whole")
                    in_data.append(offer_price.text)
                except:
                    in_data.append('Price not listed')

                try:
                    original_price = card.find_element("class name", "a-text-price")
                    in_data.append(original_price.text)
                except:
                    in_data.append('Price not listed')

                try:
                    is_in_deal = card.find_element("class name", "a-badge-label-inner")
                    in_data.append(is_in_deal.text)
                except:
                    in_data.append('not in deal')
            self.data.append(in_data)

    def save_to_excel(self, filename):
        try:
            df = pd.DataFrame(self.data)
            df.columns = ['Product Name', 'Offer Price', 'Original Price', 'Deal type']
            df.to_excel(filename, index=False)
        except Exception as e:
            print(f'Unable to write to excel: {e}')

    def close_driver(self):
        self.driver.quit()

# Example usage
if __name__ == "__main__":
    url = "https://www.amazon.in/s?i=electronics&bbn=1805560031&rh=n%3A976419031%2Cn%3A%21976420031%2Cn%3A1389401031%2Cn%3A1389432031%2Cn%3A1805560031%2Cp_6%3AA14CZOWI0VEHLG%7CA1P3OPO356Q9ZB%7CA2HIN95H5BP4BL%2Cp_89%3AApple&ref=mega_elec_s23_1_2_1_6"
    scraper = AmazonScraper(url)
    scraper.open_page()
    scraper.scrape_product_details()
    scraper.save_to_excel('mydata.xlsx')
    scraper.close_driver()
    print('Data scraped successfully! Please check mydata.xlsx.')