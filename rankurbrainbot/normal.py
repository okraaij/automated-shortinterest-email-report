from selenium import webdriver
import pandas as pd
import time, datetime, traceback, operator

start_time = time.time()

# Setup Chrome window
driver = webdriver.Chrome("C:/Users/Olivier.Kraaijeveld/Documents/Projecten/Project TurboScraper/chromedriver.exe")

driver.get('https://rankyourbrain.com/mental-math/mental-math-test-normal/play')

import operator
ops = {"+": operator.add, "-": operator.sub, "/": operator.truediv, "*": operator.mul}
opp_ops = {"+": operator.sub, "-": operator.add, "/": operator.mul, "*": operator.truediv}

# Question type 0 = filling in something after the equals sign
# Question type 1 = filling something in before the equals sign
# Question type 2 = filling something in between or after 

for i in range(0,10000):
    answer = 0
    questiontype = 0
    
    beforeanswer = driver.find_element_by_xpath("//span[@id='beforeAnswer']").text
    afteranswer = driver.find_element_by_xpath("//span[@id='afterAnswer']").text
    if "=" not in beforeanswer:
        questiontype = 1
        
    if "=" in afteranswer and beforeanswer is not "":
        questiontype = 2
    
    if questiontype == 0:
        items = beforeanswer.split(" ")
        items.remove("=")
        answer = int(ops[items[1]](int(items[0]),int(items[2])))
        
    if questiontype == 1:
        if beforeanswer == "":
            items = afteranswer.split(" ")
            items.remove("=")
            opp = opp_ops[items[0]]
            answer = int(opp(int(items[2]),int(items[1])))
            
    if questiontype == 2:
        if beforeanswer.split(" ")[0] in ops.keys():
            items = beforeanswer.split(" ") + afteranswer.split(" ")
            items.remove("=")
            opp = opp_ops[items[1]]
            answer = int(opp(int(items[2]),int(items[0])))
        else:
            items = beforeanswer.split(" ") + afteranswer.split(" ")
            items.remove("=")
            if items[1] == "-" or items[1] == "/":
                opp = ops[items[1]]
                answer = int(opp(int(items[0]),int(items[2])))
            else:
                opp = opp_ops[items[1]]
                answer = int(opp(int(items[2]),int(items[0])))
            
    element = driver.find_element_by_xpath("//input[@id='answer']")
    element.clear()
    element.send_keys(answer)    
