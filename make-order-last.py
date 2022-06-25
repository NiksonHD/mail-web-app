import requests,sys, re, bs4, string, ctypes, clipboard, webbrowser, pdfkit, os, pyqrcode, png, glob, time
from pyqrcode import QRCode
from barcode import EAN13
from configPath import *
from barcode import Code128
from barcode.writer import ImageWriter

#from selenium import map
#from selenium.webdriver.common.by import By
#from selenium.webdriver.firefox.options import Options


# options=Options()
# options.add_argument("--headless")
# browser = webdriver.Firefox(options=options)


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False
if is_admin():
	pdfFiles = glob.glob(pdfFolder + "*")
	if (len(pdfFiles)) > 1300:
		for pdf in pdfFiles:
			os.remove(pdf)
	files = glob.glob(mailsPath + "*")
	
	for f in files:
		
		
		filename = os.path.basename(f)
		
		plainFilename = filename.replace('.eml', '')
		print(plainFilename)
		
		webOrderNodate = ''
		webNumber = ''
		for word in filename.split():
			
			if len(word) == 9 and word.isnumeric():
					webOrderNodate = word
					webNumber = word
		
		text = open(f, "r")
		t = text.read()


		dateString = re.search('praktiker.bg>;(.+?)EEST',t)
		
		sapString = re.search('printable(.+?)Content-Type',t, flags=re.S)
		
		date = ''
		if dateString:
			found = dateString.group(0)
			
			d = found.replace("(EEST", "")
			date = d.replace('praktiker.bg>;', '')

		
		
		webOrder = webNumber + ' ' + date
		
		
		# search for sap numbers and quantity
		sapArray = []
		if sapString:
			rawSaps = sapString.group(0)
			
			for line in rawSaps.splitlines():
				try:
					lineArray = line.split()
					if len(lineArray[0]) < 5:
						lineArray.pop(0)
					sap = lineArray[0]
				except IndexError:
					continue
				if len(sap) == 6 and sap.isnumeric():
					sapArray.append(sap)
				#     print(word)
				try:
					sap = lineArray[1]
				except IndexError:
					continue
				quantity = lineArray[1]
				if len(quantity) < 4 and quantity.isnumeric():
				
				
					try:
						if len(sapArray[len(sapArray) -1]) > 5:
							sapArray.append(quantity)
					except IndexError:
						continue
			
		file = open(htmlPath, "w", encoding='utf-8')


		file.write('''<head>
		<meta charset="UTF-8">
		<meta name="viewport" content="width=device-width, initial-scale=1.0">
		<link rel="stylesheet" type="text/css" href="../style-web/style.css">
		<title>web-order ''' + webOrder + '''</title>
		</head>
		<div style="font-size:300%;">Web order: #''' + webOrder +'''</div>
		<div>''' + plainFilename + '''</div>
		<ol id="top">''')
		file.close()




		sapString = " ".join(sapArray)
		print(sapString)
		print(webNumber)
		qrCode = pyqrcode.create(webNumber + " " + sapString)
		qrCode.png(qrPath, scale = 6)
		divIter = 1

	

		file = open(htmlPath, "a", encoding='utf-8')

		for sapNum in sapArray:
				if len(sapNum) == 6:

					res =requests.get('https://praktiker.bg/p/' + sapNum)
					soup = bs4.BeautifulSoup(res.text, 'html.parser')
					articleName = str(soup.select('.breadcrumbs')[0]).split('<span>')[1]
					try:
						ean = soup.select('.product-code')[1]
						
					except IndexError:
						ean = '<p class="product-code">EAN: Няма в сайт.</p>'
					try:
						barcode = Code128(sapNum, writer=ImageWriter())
						barcode.save(barcodePath + sapNum)	
						pic = soup.select('.preview-media')[0]
						pic = str(pic)
						sourcePicTemp = pic.split('src="')[1]
						sourcePic = sourcePicTemp.split('"')[0]
						price = str(soup.select('.prices')[0]).split('<td colspan="3">')[0].replace('<table class="prices">','<table style="float:left;" class="prices">')
					except IndexError:
						
						continue
					

					# browser.get('https://praktiker.bg/p/'+sapNum)
					# try:
					#     elem = browser.find_element(By.CLASS_NAME, "check-availability")
					#     # print(elem)
					#     elem.click()
					#     selectButton = browser.find_element(By.ID, "pickCity_"+sapNum)
					#     selectButton.click()
					#     choice = selectButton.find_element(By.XPATH, '//*[.=\"София\"]')
					#     choice.click()
					#     town = browser.find_element(By.ID, 'pickStore_'+sapNum)
					#     town.click()
					#     choiceTown = town.find_element(By.XPATH, '//*[.=\"бул. Панчо Владигеров, 75, София, 1360\"]')
					#     choiceTown.click()
					#     onstock = browser.find_element(By.CLASS_NAME, 'available')
					#     onstock = onstock.text
					#     # print('Found <%s> element with that class name!' % (elem.tag_name))
					# except:
					#     print('Was not able to find an element with that name.')


					imgUrl = 'https://praktiker.bg'+sourcePic

					# <span>''' +onstock + '''</span>


					file.write('''<li style="font-size:23px"><div id="Div'''+str(divIter)+ ''' "><a href="https://praktiker.bg/p/
					'''+str(sapNum)+'''#globalMessages">'''+str(sapNum)+'''</a> '''+str(articleName)
					+''' ''' + str(ean) + price+'''</table><img src="'''+imgUrl+'''" width="170" height="140" style="position:absolute; margin:-70px 560px -70px 810px"> 
					<img src="'''+ barcodePath + sapNum + '''.png" width="170" height="15" style="position:absolute; margin:-10px 460px -10px 310px" >''')
					#+''' ''' + str(ean) + price+'''</table><img src="'''+imgUrl+'''" width="90" height="90"> ''')
				elif len(sapNum.strip()) < 4 and sapNum != '' and sapNum.isnumeric() :
				#<div>
					#<button id='toggle' type='button' onclick='divVisibility("Div'''+str(divIter)+ '''");'>Скрий / Покажи</button>
					orderQuantity = sapNum
					file.write('''<p>Поръчани: ''' + orderQuantity + ''' бр.</p></div>

					</li><hr>''')
					divIter+=1

		file.write('''</ol><script type=\"text/javascript\">
		function divVisibility(divId) {
			const targetDiv = document.getElementById(divId);
			// console.log(targetDiv);
			if (targetDiv.style.display !== \"none\") {
				targetDiv.style.display = \"none\";
			} else {
				targetDiv.style.display = \"block\";
			}
			// hideNonVisibleDivs();
			}
		</script>

		<div><img src="web-QR-code.png" width="170" height="170"> </div> ''')
		file.close()
		options = {
	  "enable-local-file-access": None
	}
		#set print setings to print pdf

		
		config = pdfkit.configuration(wkhtmltopdf="C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe")
		pdfkit.from_file(htmlPath, pdfFolder + webNumber + ".pdf", configuration=config, options = options )
		os.startfile(pdfFolder + webNumber + ".pdf", "print")
		# os.startfile(pdfFolder + webNumber + ".pdf")

		# open result in browser
		#webbrowser.open('file:///C:/Users/a/Music/NHD/python-scripts/web-order.html')

		# webbrowser.open('file:///C:/Users/nikso/python-scripts/web-order.html')
		
	print(sapArray)
	print(webNumber)			
	print(webOrder)
	text.close()
files = glob.glob(mailsPath + "*")
for f in files:
	os.remove(f)
barcodes = glob.glob(barcodePath + "*")
print(len(barcodes))
for b in barcodes:
	os.remove(b)	
# else:
# Re-run the program with admin rights
    # ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
