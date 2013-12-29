import urllib2
from bs4 import BeautifulSoup
import xlsxwriter

def scrap_page_data(page_num):
	
	# Excel Output Setup
	workbook = xlsxwriter.Workbook('demo.xlsx')
	worksheet = workbook.add_worksheet()
	bold = workbook.add_format({'bold': 1})
	worksheet.write('A1', 'Company Name', bold)
	worksheet.write('B1', 'Tel Num', bold)
	worksheet.write('C1', 'Address', bold)
	worksheet.write('D1', 'Categories', bold)
	worksheet.write('E1', 'Email', bold)
	worksheet.write('F1', 'Website', bold)


	url = "http://www.yellowpages.com.my/search.jsp?sfor=all&name=Plumbing+Contractors&w=&p=" + str(page_num)
	soup = BeautifulSoup(urllib2.urlopen(url).read())

	#i=0 is loop to scrap through multiple rows in 1 page
	i=0
	while i<20:
		company_name = soup('div', {'class': 'nameEL'})[i].a.text
		tel_num = soup('div', {'style': 'float:right;width:210px;text-align:right'})[i].a.text
		address = soup('div', {'class': 'addr'})[i].text
		
		#Pulls all the category listed, split up the categories, and the remove the word "and more..." and stores it back
		category_all = soup('div', {'class': 'cat'})[i].text[10:]
		#Function needs to be cleaned up further to handle more exceptions for the "and more..." problem - temporarily disabled
		#category_list = category_list_to_clean(category_all)

		detail = soup('div', {'class': 'detail1'})[i]

		#Pulls up the email encoded popup page, scraps the 2nd page and returns the email address
		email = ""
		if detail('a', {'class': 'email'}):
			email_page = detail('a', {'class': 'email'})[0].get("onclick").split("'")[1]
			email = email_page_to_scrap(email_page)

		#Pulls up the URL encoded page, evaluates the redirected page and returns the URL
		website = ""
		if detail('a', {'class': 'website'}): #[0].get("href"):
			website_encoded = detail('a', {'class': 'website'})[0].get("href")
			website = decode_website(website_encoded)
	
		
		worksheet.write(i+1, 0, company_name)
		worksheet.write(i+1, 1, tel_num)
		worksheet.write(i+1, 2, address)
		worksheet.write(i+1, 3, category_all)
		worksheet.write(i+1, 4, email)
		worksheet.write(i+1, 5, website)
				

		print company_name
		print tel_num
		print address
		print category_all
		if email != "":
			print email
		if website != "":	
			print website
		print i
		i+=1
	return


def category_list_to_clean(category_all):
	category_list = category_all.split(",")
	category_item_two = category_list[2][:-12]
	category_list.pop(2)
	category_list.insert(2,category_item_two)
	return category_list


def email_page_to_scrap(email_page):
	url = "http://www.yellowpages.com.my" + email_page
	soup = BeautifulSoup(urllib2.urlopen(url).read())
	email = soup('input', {'type': 'hidden'})[0].get("value")
	return email


def decode_website(website_encoded):
	url = "http://www.yellowpages.com.my" + website_encoded
	try:
	   website = urllib2.urlopen(url).geturl()
	except urllib2.HTTPError as e:
	   return None
	except urllib2.URLError as e:
	   return None
	except urllib2.Exception as e:
	   return None
	return website


scrap_page_data(0)
