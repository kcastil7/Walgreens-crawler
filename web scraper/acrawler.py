import urllib
from bs4 import BeautifulSoup as soup
import datetime
import xlrd
import xlwt


start = datetime.datetime.now()

file_location = "C:/Users/kevin.castillo/Desktop/WEB SCRAPER/wagMAP.xlsx"
workbook = xlrd.open_workbook(file_location)
sheet = workbook.sheet_by_index(0)
print(sheet.cell_value(1,1))
data = [[sheet.cell_value(r,c) for c in range (sheet.ncols)] for r in range(sheet.nrows)]
for rows in range (1, sheet.nrows):
	data[rows][1] = int(data[rows][1])//1
price_list = []
my_url = "https://www.walgreens.com/search/results.jsp?Ntt="
for rows in range(1,sheet.nrows):

	WIC = str(data[rows][1])
	check = True
	while check:
		try:

			req = urllib.request.Request(my_url+WIC)
			uClient = urllib.request.urlopen(req)
			page_html = uClient.read()
			uClient.close()
			page_soup = soup(page_html, "html.parser")
			#print(page_soup)
			null_div = page_soup.findAll("h4",{"class": "wag-hn-lt-55roman mt0"})
			price_div = page_soup.findAll("span",{"class": "wag-price-black wag-font-bold"})
			price_special = page_soup.findAll("span",{"class": "wag-price-red wag-text-red wag-font-bold"})
			print(len(null_div))
			print(len(price_div))
			print(len(price_special))
			print("----")
			if len(null_div) > 0:
				print("WIC:" + WIC + " not found")
				check = False
			elif len(price_div) > 0:
				#print(len(price_div))
				dollar = price_div[0].findAll("span")
				#print(len(dollar))
				cent = price_div[0].findAll("sup")
				#print(len(cent))
				print("The price of WIC " + WIC + " is $" + dollar[0].text+"."+cent[1].text)
				data[rows][5] = float(dollar[0].text+"."+cent[1].text)
				check = False
			elif len(price_special) > 0:
				#print(len(price_div))
				dollar = price_special[0].findAll("span")
				#print(len(dollar))
				cent = price_special[0].findAll("sup")
				#print(len(cent))
				print("The price of WIC " + WIC + " is $" + dollar[1].text+"."+cent[1].text)
				data[rows][5] = float(dollar[1].text+"."+cent[1].text)
				check = False

		except urllib.error.URLError as e:
			print("URL error " + e.reason + ", trying again")
		except urllib.error.HTTPError as e:
			print("HTTP error "+ e.reason + ", trying again")

print("Program took:")
print(datetime.datetime.now() - start)
workbook = xlwt.Workbook(encoding="utf-8")
sheet1=workbook.add_sheet("sheet1")

for c in range(sheet.ncols):
	for r in range(sheet.nrows):
		sheet1.write(r,c,data[r][c])


workbook.save("wagMAP.xls")
