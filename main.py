from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import pandas as pd
import numpy as np
import time
import datetime
import os

# new incognito Chrome instance
option = webdriver.ChromeOptions().add_argument(' — incognito')
driver = webdriver.Chrome(executable_path="chromedriver.exe", chrome_options=option)
driver.maximize_window()

# get the page
driver.get('https://www.ah.nl/bonus')

# set max page load time (sec)
timeout = 20

try:
    WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="accept-cookies"]')))
except TimeoutException:
    print('Timed out waiting for page to load')
    driver.quit()

# click accept cookies
driver.find_element_by_xpath('//*[@id="accept-cookies"]').click()

# fetch dates and week number
start_date = datetime.datetime.strptime(driver
                                        .find_element_by_xpath(
    "//*[@id='app']/main/div[1]/div/div/div/nav/ul/li[1]/span/span/time[1]")
                                        .get_attribute('datetime'), '%Y-%m-%dT%H:%M:%S.%fZ').date()

end_date = datetime.datetime.strptime(driver
                                      .find_element_by_xpath(
    "//*[@id='app']/main/div[1]/div/div/div/nav/ul/li[1]/span/span/time[2]")
                                      .get_attribute('datetime'), '%Y-%m-%dT%H:%M:%S.%fZ').date()

week_number = start_date.isocalendar()[1]

# scroll through page
screens = round(driver.execute_script("return document.body.scrollHeight") / 800)

for i in range(screens):
    driver.execute_script("window.scrollBy(0, 800);")
    time.sleep(0.1)

# extract prices
prices = [item.text for item in driver.find_elements_by_xpath(
    '//*[@id="app"]/main/div[2]/div/div/section[4]//div[@class="price_portrait__2WpRE"]')]
prices = [item.split('\n')[:2] for item in prices]
prices = [[x for x in item if not any(c.isalpha() for c in x)] for item in prices]

old_price = [float(item[0]) if len(item) > 1 else 'N/A' for item in prices]
new_price = [float(item[1]) if len(item) > 1 else float(item[0]) for item in prices]

perc_discount = [round(1 - new / old, 3) if isinstance(old, float) else 'N/A' for new, old in zip(new_price, old_price)]

# extract product names and descriptions
full_text = [item.text.split("\n") for item in
             driver.find_elements_by_xpath('//*[@id="app"]/main/div[2]/div/div/section[4]/div/div/article/div/div[2]')]

product_name = [item[0] for item in full_text]
product_description = [" ".join(item) if len(item) > 0 else "" for item in [item[1:] for item in full_text]]

# assign discount type
def discount_type_fun(descriptions):
    discount_type = []
    for i in range(len(descriptions)):
        if 'stapel' in descriptions[i].lower():
            discount_type.append('stapelkorting')
        elif '1 + 1' in descriptions[i].lower():
            discount_type.append('1 plus 1')
        elif '2=1' in descriptions[i].lower():
            discount_type.append('1 plus 1')
        elif '2 + 1' in descriptions[i].lower():
            discount_type.append('2 plus 1')
        elif '2 + 2' in descriptions[i].lower():
            discount_type.append('2 plus 2')
        elif '1 + 2' in descriptions[i].lower():
            discount_type.append('1 plus 2')
        elif '3 + 1' in descriptions[i].lower():
            discount_type.append('3 plus 1')
        elif '4 + 1' in descriptions[i].lower():
            discount_type.append('4 plus 1')
        elif '%' in descriptions[i].lower():
            discount_type.append('%')
        elif '2 voor' in descriptions[i].lower():
            discount_type.append('2 voor')
        elif '3 voor' in descriptions[i].lower():
            discount_type.append('3 voor')
        elif '4 voor' in descriptions[i].lower():
            discount_type.append('4 voor')
        elif '5 voor' in descriptions[i].lower():
            discount_type.append('5 voor')
        elif 'voor' in descriptions[i].lower():
            discount_type.append('voor')
        elif '2e halve prijs' in descriptions[i].lower():
            discount_type.append('2e halve prijs')
        elif '1 euro korting' in descriptions[i].lower():
            discount_type.append('1 euro korting')
        elif 'bonus' in descriptions[i].lower():
            discount_type.append('bonus')
        else:
            discount_type.append('other')
    return discount_type

discount_type = discount_type_fun([item.text for item in driver.find_elements_by_xpath(
    '//*[@id="app"]/main/div[2]/div/div/section[4]/div/div/article/div/a/div[2]//div[contains(@class, "shield_root__2R2LQ")]')])

# extract category names and lengths
category_list = [item.text for item in driver.find_elements_by_xpath(
    '//*[@id="app"]/main/div[2]/div/div/section[4]/div/header[@class="legendcard legendcard--bonus"]')]
element_list = [item.get_attribute('class') for item in
                driver.find_elements_by_xpath('//*[@id="app"]/main/div[2]/div/div/section[4]/div/*')]
category_index = [item for item in range(len(element_list)) if element_list[item] == 'legendcard legendcard--bonus'] \
                 + [len(element_list)]
category = list(np.repeat(category_list, [x2 - x1 - 1 for (x1, x2) in zip(category_index, category_index[1:])]))

# flag feature bonus folder items
bonus_feature = [item.text for item in driver.find_elements_by_xpath(
    '//*[@id="app"]/main/div[2]/div/div/section[2]/div/div/article/div/div/a')]
bonus_folder_feature = [True if item in bonus_feature else False for item in product_name]

# assign product order/placement
product_placement = range(len(product_name))

# assign brand
product_name_split = [item.lower().split(' ') for item in product_name]
brands = list(set([subitem for item in pd.read_csv('ah_brands.csv').values.tolist() for subitem in item]))
article_brand = [', '.join(list(set([brand for brand in brands if all(item in product for item in brand.lower().split(' '))])))
                 for product in product_name_split]

# compile into dataframe
df = pd.DataFrame(dict(start_date=start_date, end_date=end_date, week_number=week_number, category=category,
                       product_placement=product_placement, product_name=product_name,
                       product_description=product_description, brand=article_brand,
                       bonus_folder_feature=bonus_folder_feature, old_price=old_price, new_price=new_price,
                       perc_discount=perc_discount, discount_type=discount_type))

df.sort_values(['bonus_folder_feature', 'product_placement'], ascending=[False, True], inplace=True)

# save to excel and open
df.to_excel('AH_Bonus_week_{}.xlsx'.format(week_number))
os.startfile('AH_Bonus_week_{}.xlsx'.format(week_number))


