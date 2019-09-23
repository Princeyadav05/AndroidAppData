import play_scraper
import requests 
from bs4 import BeautifulSoup
import xlwt
import xlrd 
import wget
import urllib.request

  
loc = ("/home/prince.yadav/Desktop/Prince/applistnew.xlsx") 
  
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 

excel_file = xlwt.Workbook()
sh = excel_file.add_sheet('appdata')

  
for i in range(sheet.nrows):
	app_id = sheet.cell_value(i, 1)
	app_name = sheet.cell_value(i, 0)
	sh.write(i, 0, sheet.cell_value(i, 0))

	if app_id == '':
		print (app_name + "- SKIPPED!")
		continue
	app_details = []
	app_details = play_scraper.details(app_id)
	app_screenshots = [];
	app_screenshots = app_details['screenshots'];

	image_name = "Screenshots/" + app_name + " - 1" 
	f = open(image_name,'wb')
	f.write(urllib.request.urlopen(app_screenshots[0]).read())
	f.close()

	image_name = "Screenshots/" + app_name + " - 2" 
	f = open(image_name,'wb')
	f.write(urllib.request.urlopen(app_screenshots[1]).read())
	f.close()

	sh.write(i, 2, app_details['score'])
	sh.write(i, 3, app_details['category'])
	sh.write(i, 4, app_details['reviews'])
	sh.write(i, 5, app_details['price'])
	sh.write(i, 6, app_details['installs'])
	sh.write(i, 7, app_details['description'])
	sh.write(i, 8, app_details['url'])
	excel_file.save('names.xlsx')
	print(app_details['title'] + " - DONE!")
    # app_screenshots = []
    # app_screenshots = (app_details)['screenshots']