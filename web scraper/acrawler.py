from urllib.request import urlopen as uReq
from bs4 import BeautifulSoup as soup
import datetime
import xlrd
import xlwt


start = datetime.datetime.now()

file_location = "C:/Users/Kevin Castillo/Desktop/web scraper/wagMAP.xlsx"
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

			uClient = uReq(my_url+WIC)
			page_html = uClient.read()
			uClient.close()
		except HTTPError as e:
			if e.code == 502:
				print("Error 502, Trying again")

		else:
			page_soup = soup(page_html, "html.parser")
			#print(page_soup)
			null_div = page_soup.findAll("h4",{"class": "wag-hn-lt-55roman mt0"})
			price_div = page_soup.findAll("span",{"class": "wag-price-black wag-font-bold"})
			print(len(null_div))
			print(len(price_div))
			if len(null_div) > 0:
				print("WIC:" + WIC + " not found")
				check = False
			if len(price_div) > 0:
				#print(len(price_div))
				dollar = price_div[0].findAll("span")
				#print(len(dollar))
				cent = price_div[0].findAll("sup")
				#print(len(cent))
				print("The price of WIC " + WIC + " is $" + dollar[0].text+"."+cent[1].text)
				data[rows][5] = dollar[0].text+"."+cent[1].text
				check = False

print(start - datetime.datetime.now())
workbook = xlwt.Workbook(encoding="utf-8")
sheet1=workbook.add_sheet("sheet1")

for c in range(sheet.ncols):
	for r in range(sheet.nrows):
		sheet1.write(r,c,data[r][c])


workbook.save("wagMAP.xls")
