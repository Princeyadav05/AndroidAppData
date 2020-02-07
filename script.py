import play_scraper
import requests 
from bs4 import BeautifulSoup
import xlwt
import xlrd 
import wget
import urllib.request
from prettytable import PrettyTable


loc = ("/home/prince.yadav/Desktop/Prince/Android-App-Data-Script/applist.xlsx") 
  
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 

excel_file = xlwt.Workbook()
sh = excel_file.add_sheet('AppsData')

def setExcelHeaders():
	sh.write(0, 0, "Name")
	sh.write(0, 1, "Rating")
	sh.write(0, 2, "Category")
	sh.write(0, 3, "Number of Reviews")
	sh.write(0, 4, "Price")
	sh.write(0, 5, "Number of Downloads")
	sh.write(0, 6, "Description")
	sh.write(0, 7, "URL")

def writeDataInExcel(appsData):
	setExcelHeaders()

	## App Screenshot ##
	# app_screenshots = []
	# app_screenshots = app_details['screenshots']
	# image_name = "Screenshots/" + app_name + " - 1" 
	# f = open(image_name,'wb')
	# f.write(urllib.request.urlopen(app_screenshots[0]).read())
	# f.close()

	# image_name = "Screenshots/" + app_name + " - 2" 
	# f = open(image_name,'wb')
	# f.write(urllib.request.urlopen(app_screenshots[1]).read())
	# f.close

	print("\nWriting Data to Excel File...\n")

	for i in range(0, len(appsData)):
		sh.write(i+1, 0, appsData[i]['title'])
		sh.write(i+1, 1, appsData[i]['score'])
		sh.write(i+1, 2, appsData[i]['category'])
		sh.write(i+1, 3, appsData[i]['reviews'])
		sh.write(i+1, 4, appsData[i]['price'])
		sh.write(i+1, 5, appsData[i]['installs'])
		sh.write(i+1, 6, appsData[i]['description'])
		sh.write(i+1, 7, appsData[i]['url'])
		print(appsData[i]['title'] + " - DONE!")

	print("\nSaving Excel File...")
	excel_file.save('FinalData.xlsx')
	print("\nData Saved! Please Check FinalData.xlsx File.\n")


def displayTable(searchResults, noOfRowsToDisplay=5, detailed=False):
	table = PrettyTable()

	if detailed:
		table.field_names = ["No.", "Title", "Developer", "Category", "Rating", "Downloads"]
	else:
		table.field_names = ["No.", "Title", "Developer"]
	

	for result in range(0, noOfRowsToDisplay):
		title = searchResults[result]['title']
		developer = searchResults[result]['developer']
		category = searchResults[result]['category'] if detailed else ''
		rating = searchResults[result]['score'] if detailed else ''
		downloads = searchResults[result]['installs'] if detailed else ''

		title = title if len(title) < 50 else title[:50].rsplit(' ',1)[0] + ".."
		developer = developer if len(developer) < 30 else developer[:30].rsplit(' ', 1)[0] + ".."

		if detailed:
			table.add_row([result+1, title, developer, category[0], rating, downloads])
		else:
			table.add_row([result+1, title, developer])

	table.align['Title'] = "l"
	table.align['Downloads'] = "r"
	print("\nSearch Results for " + searchResults[0]['title'] + ": \n")
	print(table.get_string())

def getAppsData(sheet, detailed):
	app_details = []
	print(detailed)
	for i in range(1, sheet.nrows):
		# app_id = sheet.cell_value(i, 1)
		app_name = sheet.cell_value(i, 0)
		print("\nSearching Apps for " + app_name + "...")

		if detailed:
			searchResults = play_scraper.search(app_name, page=1, detailed=detailed)
			displayTable(searchResults, noOfRowsToDisplay=10, detailed=detailed)

			selectedApp = int((input("\nPlease Select your App: ")))
			print("\nSaved Data for " + searchResults[selectedApp - 1]['title'] + ".")

			app_details.append(searchResults[selectedApp - 1])

		else:
			searchResults =  play_scraper.search(app_name, page=1)
			displayTable(searchResults, noOfRowsToDisplay=10)

			selectedApp = int((input("\nPlease Select your App: ")))
			print("\nSaved Data for " + searchResults[selectedApp - 1]['title'] + ".")

			appId = searchResults[selectedApp - 1]['app_id']
			app_details.append(play_scraper.details(appId))
		
	return app_details

	# sh.write(i, 0, sheet.cell_value(i, 0))
# 
	# if app_id == '':
		# print (app_name + "- SKIPPED!")
		# continue

detailed = True
appsData = getAppsData(sheet, detailed)
# print(len(appsData))
writeDataInExcel(appsData)