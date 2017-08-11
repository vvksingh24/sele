# importing libraries
import selenium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import xlwt
import time
import os
import urllib.request
from PIL import Image

#for downloading image

opener=urllib.request.build_opener()
opener.addheaders=[('User-Agent','Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/36.0.1941.0 Safari/537.36')]
urllib.request.install_opener(opener)

#providing input

cwd=os.getcwd()
path=(os.path.join(cwd,"images"))
if not os.path.exists(path):
	os.makedirs(path)
#webdriver

driver= webdriver.Chrome()
driver.maximize_window()
driver.get('https://www.tmdn.org/tmview/welcome')
elem=driver.find_element_by_partial_link_text('Advanced')
elem.click()
trad=driver.find_element_by_name('TrademarkName')
trad.clear()
name="vivekananda"
trad.send_keys(name)
search=driver.find_element_by_id('SearchCopy')
search.click()

#workbook

wb=xlwt.Workbook(encoding='utf-8')
ws=wb.add_sheet('test_sheet')
ws.write(0,0,"graphic representations")
ws.write(0,1,"Trademark name")
ws.write(0,2,"Trademark office")
ws.write(0,3,"Designated Territory")
ws.write(0,4,"Application No")
ws.write(0,5,"Registration No")
ws.write(0,6,"Trademark Status")
ws.write(0,7,"Nice class")
ws.write(0,8,"Applicant name")
ws.write(0,9,"Application Date")
ws.write(0,10,"Trademark Type")
ws.write(0,11,"Registration Date")
ws.write(0,12,"Seniority claimed")
count=1

#getting data

while True:
	time.sleep(10) #waiting for data
	tab=driver.find_element_by_xpath('//table[@id="grid"]')
	tdata=tab.find_elements_by_tag_name('tr')
	for t in range(1,len(tdata)):
		a=[] #for excluding non visible data
		path=os.path.join(os.getcwd(),"images")
		i=tdata[t].find_elements_by_tag_name('td')
		for j in range(4,len(i)):
			if "display: none" in i[j].get_attribute('style'):
				continue
			else:
				a.append(i[j].text)

		good=driver.find_element_by_link_text(i[7].text)
		good.click()
		frame = driver.find_element_by_tag_name('iframe')

		for j in range(len(a)):
			ws.write(count,j+1,a[j])

		try:
			img=i[3].find_element_by_tag_name('img')
			src=img.get_attribute("src")
			# print (src)
			# print(type(i[7].text))
			path=(os.path.join(path,i[7].text+".jpeg"))
			temp=path.replace("jpeg","bmp")
			# print (temp)
			path.replace("\\","\\\\")
			# print (path)
			image=urllib.request.urlretrieve(src,path)
			# img=Image.open(image)
			# r,g,b,a = img.split()
			# img = Image.merge("RGB", (r, g, b))
			# img.save('imagetoadd.bmp')
			# xlwt.insert_bitmap('imagetoadd.bmp', count, 0)
			Image.open(path).convert("RGB").save(temp)
			ws.insert_bitmap(temp,count,0)
			print ("partial")
			# formula = 'HYPERLINK("{}", "{}")'.format(path, path)
			# ws.write(count, 0, xlwt.Formula(formula))

		except Exception as e:
			ws.write(count,0,"na")
		count+=1

	n=driver.find_element_by_xpath('//*[@id="next_grid_pager"]')
	if n.get_attribute("class")=="ui-pg-button ui-corner-all ui-state-disabled":
		break
	else:
		n.click()

wb.save(name+'.xls')
driver.close()