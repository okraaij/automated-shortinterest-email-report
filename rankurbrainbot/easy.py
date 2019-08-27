from selenium import webdriver
import pandas as pd
import time, datetime, traceback, operator

start_time = time.time()

# Setup Chrome window
driver = webdriver.Chrome("C:/Users/Olivier.Kraaijeveld/Documents/Projecten/Project TurboScraper/chromedriver.exe")

driver.get('https://rankyourbrain.com/mental-math/mental-math-test-easy/play')

import operator
ops = {"+": operator.add, "-": operator.sub, "/": operator.truediv, "*": operator.mul}

for i in range(0,10000):
    answer = 0
    
    beforeanswer = driver.find_element_by_xpath("//span[@id='beforeAnswer']").text
    afteranswer = driver.find_element_by_xpath("//span[@id='afterAnswer']").text

    if questiontype == 0:
        items = beforeanswer.split(" ")
        items.remove("=")
        answer = int(ops[items[1]](int(items[0]),int(items[2])))
        element = driver.find_element_by_xpath("//input[@id='answer']")
        element.send_keys(answer) 
