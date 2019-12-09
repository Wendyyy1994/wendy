from bs4 import BeautifulSoup
import requests
import pandas as pd 
import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
import re,time,random


data = pd.DataFrame(pd.read_excel(r'C:\Users\lenovo\Desktop\homedepot data.xlsx'))
[data['title'], data['price'], data['if dis'], data['out of stock'], data['review count'], data['review rating'], data['pic count'], data['review']] = [None]*8

#url_list = []#拼接链接
option = webdriver.ChromeOptions()
option.add_argument('--start-maximized')#窗口最大化
option.add_argument('--disable-infobars')# 禁用浏览器正在被自动化程序控制的提示
option.add_argument('--incognito')# 隐身模式，无痕浏览
option.add_argument('headless')#不显示可视化界面
cc = webdriver.Chrome(chrome_options=option)
#cc = webdriver.Chrome()
#for j in range (len(data)):
    #url_list.append('https://www.homedepot.com/p/' + str(data['SKU'][j]))

for i in range(0,len(data)):#len(data)
    cc.get(data['URL'][i])
    time.sleep(2)    
    try:
        title = cc.find_element_by_tag_name('h1').text
    except:
        title = 'off-line'    
    try:
        dis = cc.find_element_by_xpath('//*[@class="product-title__wrapper"]/h2/span[@class="u__text--danger"]').text
    except:
        try:
            dis = cc.find_element_by_xpath('//*[@class="alert-oos-hover-wrapper"]/div/span').text
        except:
            dis = ''
    try:
        outofstock = cc.find_element_by_xpath('//*[@class="buybelt__box shipping-box"]/div').text.replace('Store','').replace('\n','')
    except:
        outofstock = ''
        
    try:
        p = cc.find_elements_by_xpath('//*[@class="price-format__large price-format__main-price"]/span')
        price = p[0].text +p[1].text +'.'+p[2].text
    except:
        try:
            p = cc.find_element_by_xpath('//*[@class="price__wrapper"]/span') 
            price = '$'+ p.get_attribute('content') 
        except:
            p = ''
            price = ''
        
    try:
        group = cc.find_element_by_xpath('//*[@class="review-rating"]/div/div[@class="col__12-12 col__4-12--xs"]/ul') 
        reviewcount = group.find_elements_by_tag_name('li')[0].text
        reviewrating = group.find_elements_by_tag_name('li')[2].text
    except:
        group = ''
        reviewcount = ''
        reviewrating = ''

    try:
        p = len(cc.find_elements_by_xpath('//*[@class="mediagallery__thumbnails"]/div')) 
        if p != 0:
            piccount = p
        else:
            piccount = len(cc.find_elements_by_xpath('//*[@class="media__thumbnail"]'))  
    except:
        piccount = ''   
    
    try: 
        cc.find_element_by_xpath('//*[@class="product-details__review-count"]').click()
        time.sleep(2)
        try:
            review = ''
            page = cc.find_elements_by_xpath('//*[@class="hd-pagination__link "]')[-2].text
            for w in range(0,int(page)):
                review_body = cc.find_elements_by_xpath('//*[@class="review-content-body"]')
                for r in range(len(review_body)):
                    review += '\n'.join([str(r+1)+ '. '+ review_body[r].text])
                    
                cc.find_elements_by_xpath('//*[@class="hd-pagination__item hd-pagination__button"]/a')[-1].click()
                time.sleep(2)
        except:
                review = ''
                review_body = cc.find_elements_by_xpath('//*[@class="review-content-body"]')
                for r in range(len(review_body)):
                    review += '\n'.join([str(r+1)+ '. '+ review_body[r].text])
    except:
        review = ''
                    
                      
    data['title'][i], data['price'][i], data['if dis'][i], data['out of stock'][i], data['review count'][i], data['review rating'][i], data['pic count'][i], data['review'][i] = title, price, dis, outofstock, reviewcount, reviewrating, piccount, review 
        
    print('\t'.join([time.strftime('%H:%M:%S',time.localtime()), str(i), str(data['SKU'][i]), str(len(data))]))

cc.quit()
data.to_excel('homedepot anotherhalfonlinedata.xls')
    


#%%homedepot com chair 数据 
import pandas as pd 
import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
import re,time,random

output = [['title'], ['USKU'], ['price'], ['statue'], ['review count'], ['brand'], ['page'], ['reviewrating'],['main SKU']] 

option = webdriver.ChromeOptions()
option.add_argument('--start-maximized')#窗口最大化
option.add_argument('--disable-infobars')# 禁用浏览器正在被自动化程序控制的提示
option.add_argument('--incognito')# 隐身模式，无痕浏览
option.add_argument('headless')#不显示可视化界面
cc = webdriver.Chrome(chrome_options=option) 
cc = webdriver.Chrome()

url = 'https://www.homedepot.com/b/Furniture-Living-Room-Furniture-Chairs-Recliners/N-5yc1vZcf7s?sortby=topsellers&sortorder=desc'
cc.get(url)
time.sleep(2)
page = cc.find_elements_by_xpath('//*[@class="hd-pagination__link "]')[-2].text
for w in range(0,7): #int(page)  
    gs = cc.find_elements_by_xpath('//*[@class="product-result__wrapped-results"]/div/div')    
    for j in range (len(gs)):
        #颜色group
        try:
            group = cc.find_elements_by_xpath('//*[@class="productVariant__wrapper"]/div')[j]
            group.find_element_by_tag_name('div')
            s = group.find_elements_by_tag_name('div')
            
            for f in range(len(s)):

                s[f].find_element_by_tag_name('button').click()
                time.sleep(1)
                gs1 = cc.find_elements_by_xpath('//*[@class="product-result__wrapped-results"]/div/div') 
                sku = gs1[j]
                output[1].append(sku.get_attribute('data-itemid'))
                output[6].append(w+1)
                
                #title group               
                try:
                    group1 = cc.find_elements_by_xpath('//*[@class="product-pod__title__product"]')[j] 
                    t = group1.find_elements_by_tag_name('span')
                    if len(t) == 2:
                        output[0].append(t[1].text)
                    else:
                        output[0].append(t[0].text)
                except:
                    group1 =''
                    t = ''
                        
                #brand group
                try:
                    group5 = cc.find_elements_by_xpath('//*[@class="product-pod--padding product-pod--ie-fix"]/div[@class="product-pod__title"]')[j]
                    output[5].append(group5.find_element_by_tag_name('p').text)
                except:
                    output[5].append('null')
                            
                #review group              
                try:
                    group2 = cc.find_elements_by_xpath('//*[@class="product-pod__rating__count"]')[j]
                    output[4].append(group2.text.replace('(','').replace(')',''))
                except:
                    output[4].append('null')
                    
                #statue group     
                try:           
                    group4 = cc.find_elements_by_xpath('//*[@class="product-pod__fulfillment--shipping"]/div/span')[j]
                    output[3].append(group4.text)
                except:
                    output[3].append('null')
                # price group
                try:
                    g = cc.find_elements_by_xpath('//*[@class="product-pod__pricing"]/div/div')[j]    
                    output[2].append(g.find_elements_by_tag_name('span')[0].text + g.find_elements_by_tag_name('span')[1].text +'.'+g.find_elements_by_tag_name('span')[2].text)
                except:
                    if group4.text == 'Delivery unavailable':
                        output[2].append('null') 
                    
                #review rating%
                try:
                    rat = cc.find_elements_by_xpath('//*[@class="grid product-pod__rating"]/div/span')[j]  
                    output[7].append(rat.get_attribute('style').replace('width: ',''))
                except:
                    rat = ''
                    output[7].append('null') 
                print('\t'.join([time.strftime('%H:%M:%S',time.localtime()),' 当前sku:',str(output[1][-1]),' 当前页数:',str(w+1)]))
                            
        except:
            sku = gs[j]
            output[1].append(sku.get_attribute('data-itemid'))
            output[6].append(w+1)
            #title group               
            try:
                group1 = cc.find_elements_by_xpath('//*[@class="product-pod__title__product"]')[j] 
                t = group1.find_elements_by_tag_name('span')
                if len(t) == 2:
                    output[0].append(t[1].text)
                else:
                    output[0].append(t[0].text)
            except:
                group1 =''
                t = ''
                output[0].append('null')
            #brand group
            try:
                group5 = cc.find_elements_by_xpath('//*[@class="product-pod--padding product-pod--ie-fix"]/div[@class="product-pod__title"]')[j]
                output[5].append(group5.find_element_by_tag_name('p').text)
            except:
                output[5].append('null')
                            
            #review group              
            try:
                group2 = cc.find_elements_by_xpath('//*[@class="product-pod__rating__count"]')[j]
                output[4].append(group2.text.replace('(','').replace(')',''))
            except:
                output[4].append('null')
                
            
            #statue group     
            try:           
                group4 = cc.find_elements_by_xpath('//*[@class="product-pod__fulfillment--shipping"]/div/span')[1]
                output[3].append(group4.text)
            except:
                output[3].append('null')
            # price group
            try:
                g = cc.find_elements_by_xpath('//*[@class="product-pod__pricing"]/div/div')[j]    
                output[2].append(g.find_elements_by_tag_name('span')[0].text + g.find_elements_by_tag_name('span')[1].text +'.'+g.find_elements_by_tag_name('span')[2].text)
            except:
                if group4.text == 'Delivery unavailable':
                    output[2].append('null')
            
            #review rating%
            try:
                rat = cc.find_elements_by_xpath('//*[@class="grid product-pod__rating"]/div/span')[j]  
                output[7].append(rat.get_attribute('style').replace('width: ',''))
            except:
                rat = ''
                output[7].append('null') 

            print('\t'.join([time.strftime('%H:%M:%S',time.localtime()),' 当前sku:',str(output[1][-1]),' 当前页数:',str(w+1)]))

    cc.find_elements_by_xpath('//*[@class="hd-pagination__item hd-pagination__button"]/a')[-1].click()
    time.sleep(2)
cc.quit()

x = {'title':output[0][1:],
    'USKU':output[1][1:],
    'price':output[2][1:],
    'statue':output[3][1:],
    'review count':output[4][1:],
    'brand':output[5][1:],
    'now page':output[6][1:],
    'reviewrating':output[7][1:]}

data = pd.DataFrame(x)
data.to_excel('homedepot Recliner data.xls')
                
                    
            
            
        
        
        
       

