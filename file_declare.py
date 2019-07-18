import os
import csv
import uno
from datetime import date
from com.sun.star.uno import RuntimeException

doc = XSCRIPTCONTEXT.getDocument()

url = doc.URL
dirname = os.path.dirname(url)

filename_lst = []
data = []
thead_lst = ['統編','','總餘額','淨值比例','交易行為類別','交易金額','交易行為涉及之標的種類及金額',
	'內容說明','總餘額','淨值比例','交易金額','','統編長度']
	
transaction_name_dict = {
	3:'授信（不含短期票券之保證或背書等）',
	4:'短期票券之保證或背書等',
	5:'票券或債券之附賣回交易', 
	6:'投資或購買左列對象為發行人之有價證券',
	7:'衍生性金融商品交易',
	8:'其他經主管機關規定之交易'
}

transaction_detail_dict = {
	3:{9:'無擔保金額',10:'有擔保金額'},
	
	4:{11:'無擔保金額',12:'有擔保金額'},
	
	5:{13:'票券-有保證機構',14:'公債-有保證機構',15:'公司債券及金融債-有保證機構',16:'其他有價證券-有保證機構',
	   17:'票券-無保證機構',18:'公債-無保證機構',19:'公司債券及金融債-無保證機構',20:'其他有價證券-無保證機構'},
	
	6:{21:'票券-有保證機構',22:'公債-有保證機構',23:'公司債券及金融債-有保證機構',24:'股票-有保證機構',
	   25:'受益憑證-有保證機構',26:'證券化商品-有保證機構',27:'其他有價證券-有保證機構',28:'票券-無保證機構',
	   29:'公債-無保證機構',30:'公司債券及金融債-無保證機構',31:'股票-無保證機構',32:'受益憑證-無保證機構',
	   33:'證券化商品-無保證機構',34:'其他有價證券-無保證機構'},
	
	7:{35:'證券衍生性商品交易',36:'利率衍生性商品交易',37:'資產交換交易',38:'結構型商品交易',39:'股權衍生性商品交易',
	   40:'信用衍生性商品交易',41:'外匯衍生性商品交易',42:'其他衍生性商品交易'},
	
	8:{43:'未分類'}
}

def filename_handle():
	
	filename_lst.clear()
	path = uno.fileUrlToSystemPath(dirname)
	
	files = os.listdir(path)
	for file in files:
		if '.csv' in str(file) and 'lock' not in str(file):
			filename_lst.append(file)
	
	filename_lst.insert(len(filename_lst)-1, filename_lst.pop(0))
	
	for i in range(len(filename_lst)):
		filename_lst[i] = os.path.join(path, filename_lst[i])

def file_reader():
	
	filename_handle()
	data.clear()
	
	for i in range(len(filename_lst)):
		if 'candp' in str(filename_lst[i]):
			thead_lst[1] = '同一自然人或法人'
		if 'p2p' in str(filename_lst[i]):
			thead_lst[1] = '同一自然人與其配偶、二親等以內之血親，以本人或配偶為負責人之企業'
		if 'c2c' in str(filename_lst[i]):
			thead_lst[1] = '同一集團法人關係企業'
		data.append(thead_lst.copy())
		
		with open(filename_lst[i], mode='r') as source_file:
			csv_file = csv.reader(source_file, delimiter=',')
			for i, row in enumerate(csv_file):
				if i == 0:
					continue
				data.append(row[3:])
				
def data_handle():
	
	for i, row in enumerate(data):
		if len(row) == len(thead_lst):
			continue
		for j, column in enumerate(row):
			if j > 1 and j < len(row)-1 and row[j].isdigit():
				row[j] = int(column)
			if j == len(row)-1:
				row[j] = float(column)
				
def refresh_match_startrow():
	sheet = doc.Sheets[1]
	cursor = doc.Sheets[0].createCursor()
	cursor.gotoEndOfUsedArea(False)
	lastrow = cursor.getRangeAddress().EndRow + 1
	
	refresh(sheet, 'F4', '=MATCH(D4;$申報資料.$B$1:$B${};0)'.format(lastrow))
	refresh(sheet, 'F5', '=MATCH(D5;$申報資料.$B$1:$B${};0)'.format(lastrow))
	refresh(sheet, 'F6', '=MATCH(D6;$申報資料.$B$1:$B${};0)'.format(lastrow))
	
	refresh(sheet, 'E4', '=COUNTA($申報資料.B2:B{})'.format(
		int(sheet.getCellRangeByName('F5').getValue()-1)))
	refresh(sheet, 'E5', '=COUNTA($申報資料.B{}:B{})'.format(
		int(sheet.getCellRangeByName('F5').getValue()+1),
		int(sheet.getCellRangeByName('F6').getValue()-1)))
	refresh(sheet, 'E6', '=COUNTA($申報資料.B{}:B{})'.format(
		int(sheet.getCellRangeByName('F6').getValue()+1), lastrow))
	
def refresh(sheet, name, formula):
	cell = sheet.getCellRangeByName(name)
	cell.setFormula(formula)
	
def id_handle(_id):
	
	for ch in _id:
		if ch.isalpha(): return _id
	
	if len(_id) <= 8:
		tmp = int(_id)
		return '{:08}'.format(tmp)
				
def data_declare():

	file_reader()
	data_handle()
	
	locale = doc.CharLocale
	numbers = doc.NumberFormats
	try:
		n = numbers.addNew('#,##0',locale)
	except RuntimeException:
		n = numbers.queryKey('#,##0',locale,False)

	sheet = doc.Sheets[0]
	sheet.Name = '申報資料'
		
	for i,row in enumerate(data):
		cursor = sheet.createCursor()
		cursor.gotoEndOfUsedArea(False)
		if len(row) == len(thead_lst):
			if i == 0:
				lastrow = 0
			else:
				lastrow = cursor.getRangeAddress().EndRow + 1
			cell = sheet.getCellRangeByPosition(0,lastrow,len(row)-1,lastrow)
			cell.CellBackColor = 0x92d050
			cell.setDataArray((row,))
			continue
		
		lastrow = cursor.getRangeAddress().EndRow + 1
		oRange = sheet.getCellRangeByPosition(0,lastrow,12,i*6)
		
		id = id_handle(row[0].strip())
		cell = oRange.getCellByPosition(0,0)
		cell.String = str(id)
		
		cell = oRange.getCellByPosition(1,0)
		cell.String = row[1]
		
		cell = oRange.getCellByPosition(2,0)
		cell.NumberFormat = n
		cell.Value = row[2]
		
		cell = oRange.getCellByPosition(3,0)
		cell.Value = row[-1]
		
		for j, key in enumerate(transaction_name_dict.keys()):
			cell = oRange.getCellByPosition(0,j)
			cell.CellBackColor = 0x99ccff
		
			name = oRange.getCellByPosition(4,j)
			name.String = transaction_name_dict[key]
			if str(row[key]).strip() == str(0): continue
			sheet.getCellRangeByName('N2').String = row[key]
			money = oRange.getCellByPosition(5,j)
			money.Value = row[key]
			money.NumberFormat = n
			detail = oRange.getCellByPosition(6,j)
			for idx in transaction_detail_dict[key].keys():
				if str(row[idx]).strip() != str(0):
					detail.String += '{:}:{:,}\n'.format(transaction_detail_dict.get(key).get(idx),int(row[idx]))
			detail.String = detail.getString().rstrip('\n')
			
			cell = oRange.getCellByPosition(10,j)
			cell.setFormula('=ROUND(F{0}/1000;0)'.format(cell.getCellAddress().Row+1))
			cell.NumberFormat = n
			
		cell = oRange.getCellByPosition(8,0)
		cell.setFormula('=ROUND(C{0}/1000;0)'.format(lastrow+1))
		cell.NumberFormat = n
		
		cell = oRange.getCellByPosition(9,0)
		cell.Value = row[-1] * 100
		
		cell = oRange.getCellByPosition(12,0)
		cell.setFormula('=LEN(A{0})'.format(lastrow + 1))
		
	refresh_match_startrow()
		
g_exportedScripts = data_declare,