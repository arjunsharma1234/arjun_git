import sys
reload(sys)
sys.setdefaultencoding('utf8')
import pandas as pd
import re
from PyQt4.QtCore import *
from PyQt4.QtGui import *
from PyQt4 import *
import time
import pymysql.cursors
import pymysql
import os.path
import csv
import datetime
from PyQt4.QtWebKit import QWebView
import xlsxwriter

def window():
   
   app = QApplication(sys.argv)
   win = QWidget()
   grid = QGridLayout()
   
   win.setLayout(grid)
   win.setGeometry(500,500,1270,1270)
   win.setWindowTitle("IB Price Comparison 1.0.1")
   
   win1 = QWidget()
   win1.setLayout(grid)
   win1.setGeometry(600,600,900,500)
   win1.setWindowTitle("SKU Tracker")
   
   win2 = QWidget()
   win2.setLayout(grid)
   win2.setGeometry(400,400,700,400)
   win2.setWindowTitle("Detail")
   
   win3 = QWidget()
   win3.setLayout(grid)
   win3.setGeometry(400,400,700,400)
   win3.setWindowTitle("Detail")
   
   win4 = QWidget()
   win4.setLayout(grid)
   win4.setGeometry(400,400,1100,650)
   win4.setWindowTitle("Sku Stats")
   
   
   msg = QMessageBox()
   msg.setIcon(QMessageBox.Information)
   msg.setWindowTitle("Information")
   
   table = QtGui.QTableWidget(win)
   table.setFixedWidth(1250)
   table.setFixedHeight(680)
   table.setColumnCount(11)
   table.setHorizontalHeaderLabels(QString("IB SKU;IB Category;IB Status;IB Price;Moglix Status;Moglix Price;Snapdeal Status;Snapdeal Price;Amazon Status;Amazon Price;Links").split(";"))
   table.setColumnWidth(0,206)
   for i in range (1,12):
		table.setColumnWidth(i,118)
   table.move(10,30)
   
   table1 = QtGui.QTableWidget(win1)
   table1.setFixedWidth(870)
   table1.setFixedHeight(470)
   table1.setColumnCount(11)
   table1.setHorizontalHeaderLabels(QString("IB SKU;IB Category;IB Status;IB Price;Moglix Status;Moglix Price;Snapdeal Status;Snapdeal Price;Amazon Status;Amazon Price;Date").split(";"))
   table1.setColumnWidth(0,100)
   for i in range (1,12):
		table1.setColumnWidth(i,80)
   table1.move(15,13)
   
   table2 = QtGui.QTableWidget(win3)
   table2.setFixedWidth(870)
   table2.setFixedHeight(470)
   table2.setColumnCount(11)
   table2.setHorizontalHeaderLabels(QString("IB SKU;IB Category;IB Status;IB Price;Moglix Status;Moglix Price;Snapdeal Status;Snapdeal Price;Amazon Status;Amazon Price;Date").split(";"))
   table2.setColumnWidth(0,100)
   for i in range (1,12):
		table2.setColumnWidth(i,80)
   table2.move(15,13)
   
   table3 = QtGui.QTableWidget(win4)
   table3.setFixedWidth(1080)
   table3.setFixedHeight(630)
   table3.move(10,30)
   
   win.show()
   
   filter_col = QLineEdit(win)
   filter_col.setPlaceholderText("Filter SKU")
   filter_col.setFixedWidth(200)
   filter_col.move(220,1)
   filter_col.show()
   
   def on_filter():
	   
	   QtGui.QApplication.setOverrideCursor(QtCore.Qt.WaitCursor)
	   filter_text = filter_col.text()
	   allrows = table.rowCount()
	   table2.setRowCount(allrows)
	   for row in range(0,allrows):
			table2.showRow(row)
			item = table.item(row,0)
			text = str(item.text())
			if str(filter_text) in text:
				for i in range(10):
					table2.setItem(row,i,QTableWidgetItem(table.item(row,i)))
				table2.setItem(row,10,QTableWidgetItem(str(cb.currentText())))
			else:
				table2.hideRow(row)
	   QtGui.QApplication.restoreOverrideCursor()
	   table2.show()
	   win3.show()
			
   
   filter_but = QPushButton("Filter",win)
   filter_but.clicked.connect(on_filter)
   filter_but.move(430,1)
   filter_but.setFixedWidth(100)
   filter_but.setStyleSheet("background-color:lightblue")
   filter_but.show()
   
   cb = QtGui.QComboBox(win)
   cb.move(1050,2)
   cb.setFixedWidth(120)
   
   try:
		db = pymysql.connect("35.154.31.255","crawl","St@nD@#d","scrap")
		sql = "SELECT DISTINCT(Date) FROM Ib_Price"
		dates = pd.read_sql(sql,db)
		for date in dates['Date'].tolist():
			cb.addItem(date)
		cb.show()
		
   except:
	   msg.setText("Connection Problem")
	   msg.show()
	   win.close()
	   QtGui.QApplication.restoreOverrideCursor()
   
   def fetch():
		   try:
				QtGui.QApplication.setOverrideCursor(QtCore.Qt.WaitCursor)
				db = pymysql.connect("35.154.31.255","crawl","St@nD@#d","scrap")
				abc = str(cb.currentText())
				sql = """SELECT t1.`IB Sku`,
					t1.`IB Category`,
					t1.`IB Status`,
					t1.`IB Price`,
					t1.`Moglix Status`,
					t1.`Moglix Price`,
					t1.`Snapdeal Status`,
					t1.`Snapdeal Price`,
					t1.`Amazon Status`,
					t1.`Amazon Price`,
					t1.`Date`,
					CONCAT(t2.`Moglix URL`," - ",
					t2.`Snapdeal URL`," - ",
					t2.`Amazon URL`) AS `Links`
					FROM `Ib_Price` t1
					INNER JOIN sku_links t2
					ON t1.`IB Sku` = t2.Sku
					WHERE t1.`Date` ="""+"'"+str(cb.currentText())+"'"
		   
				data1 = pd.read_sql(sql,db)
				table.setShowGrid(False)
		   except:
				msg.setText("Connection Problem")
				msg.show()
				QtGui.QApplication.restoreOverrideCursor()
		   
		   
		   def showitem(row,column):
			   
			   tx = QTextEdit(win2)
			   tx.setFixedWidth(650)
			   tx.setFixedHeight(325)
			   tx.move(20,20)
			   item = table.item(row,column) 
			   val = item.text()
			   values = val.split(' - ')
			   
			   final_html = ' '
			   for i,value in enumerate(values):
				   if i==0:
						vel = '<a href='+str(value)+'>'+str(value)+'</a>'
				   if i==1:
					   vel = '<a href='+str(value)+'>'+str(value)+'</a>'
				   if i==2:
					   vel = '<a href='+str(value)+'>'+str(value)+'</a>'
				   final_html = final_html + vel + '<br><br>'
			   tx.setHtml(final_html)
			   
			   browser = QWebView()
			   
			   def show():
				  
				  QtGui.QApplication.setOverrideCursor(QtCore.Qt.WaitCursor) 
				  cursor = tx.textCursor()
				  textSelected = cursor.selectedText()
				  browser.load(QUrl(str(textSelected)))
				  browser.show()
				  QtGui.QApplication.restoreOverrideCursor()
				  
			   
			   show_but = QPushButton("Show",win2)
			   show_but.clicked.connect(show)
			   show_but.setStyleSheet("background-color:skyblue")
			   show_but.move(20,355)
			   show_but.setFixedWidth(650)
			   
			   win2.show()
		   
		   
		   table.cellDoubleClicked.connect(showitem)
		   table.setRowCount(len(data1))
		   for i in range (len(data1)):
				table.setRowHeight(i,50)
			
		   for index, row in data1.iterrows():
				if (str(row["Moglix Price"])) == "nan":
					row["Moglix Price"] = "-"
				if (str(row["Amazon Price"])) == "nan":
					row["Amazon Price"] = "-"
				if (str(row["Snapdeal Price"])) == "nan":
					row["Snapdeal Price"] = "-"
				table.setItem(index,0,QTableWidgetItem(str(row["IB Sku"])))
				table.setItem(index,1,QTableWidgetItem(str(row["IB Category"])))
				table.setItem(index,2,QTableWidgetItem(str(row["IB Status"])))
				table.setItem(index,3,QTableWidgetItem(str(row["IB Price"])))
				table.setItem(index,4,QTableWidgetItem(str(row["Moglix Status"])))
				table.setItem(index,5,QTableWidgetItem(str(row["Moglix Price"])))
				table.setItem(index,6,QTableWidgetItem(str(row["Snapdeal Status"])))
				table.setItem(index,7,QTableWidgetItem(str(row["Snapdeal Price"])))
				table.setItem(index,8,QTableWidgetItem(str(row["Amazon Status"])))
				table.setItem(index,9,QTableWidgetItem(str(row["Amazon Price"])))
				table.setItem(index,10,QTableWidgetItem(str(row["Links"])))
				if ((row["IB Price"] < row["Moglix Price"]) and (row["IB Price"] < row["Snapdeal Price"]) and (row["IB Price"] < row["Amazon Price"])):
					for i in range(11):
						table.item(index,i).setBackground(QtGui.QColor("lightgreen"))
				else:
					for i in range(11):
						table.item(index,i).setBackground(QtGui.QColor("Red"))
				if ((row["Moglix Price"] == "-") and (row["Snapdeal Price"] == "-") and (row["Amazon Price"] == "-")):
					for i in range(11):
						table.item(index,i).setBackground(QtGui.QColor("yellow"))
		   table.show()
		   QtGui.QApplication.restoreOverrideCursor()
   
   search = QtGui.QPushButton("Fetch",win)
   search.setStyleSheet("background-color:skyblue")
   search.clicked.connect(fetch)
   search.setFixedWidth(70)
   search.move(1175,1)
   search.show()
   
   def export():
	   QtGui.QApplication.setOverrideCursor(QtCore.Qt.WaitCursor)
	   allrows = table.rowCount()
	   if allrows==0:
			msg.setText("Blank Table")
			msg.show()
			QtGui.QApplication.restoreOverrideCursor()
	   else:
		 try:
			npt = datetime.datetime.now()
			npt = re.sub('\s+|:|\.','-',str(npt))
			for row in range(0,allrows):
				item = table.item(row,0).text()
				item1 = table.item(row,1).text()
				item2 = table.item(row,2).text()
				item3 = table.item(row,3).text()
				item4 = table.item(row,4).text()
				item5 = table.item(row,5).text()
				item6 = table.item(row,6).text()
				item7 = table.item(row,7).text()
				item8 = table.item(row,8).text()
				item9 = table.item(row,9).text()
				item10 = table.item(row,10).text()
				file_exists = os.path.isfile('IB_Pricing_Result_'+str(npt)+'.csv')
				with open('IB_Pricing_Result_'+str(npt)+'.csv', 'a') as csvfile:
					fieldnames = ['IB SKU','IB Category','IB Status','IB Price','Moglix Status','Moglix Price',
					'Snapdeal Status','Snapdeal Price','Amazon Status','Amazon Price','Links','Date']
					writer = csv.DictWriter(csvfile, fieldnames=fieldnames)	
					if not file_exists:	
						writer.writeheader()
					writer.writerow({'IB SKU':item,'IB Category':item1,'IB Status':item2,'IB Price':item3,
					'Moglix Status':item4,'Moglix Price':item5,
					'Snapdeal Status':item6,'Snapdeal Price':item7,'Amazon Status':item8,'Amazon Price':item9,'Links':item10,'Date':str(cb.currentText())})
				msg.setText("Data exported to CSV")
				msg.show()
			QtGui.QApplication.restoreOverrideCursor()	
		 except:
			msg.setText("Error Occured")
			msg.show()
			QtGui.QApplication.restoreOverrideCursor()
				
   
   export_but = QPushButton("Export to csv",win)
   export_but.clicked.connect(export)
   export_but.setStyleSheet("background-color:lightgreen")
   export_but.move(890,1)
   export_but.setFixedWidth(120)
   export_but.show()
   
   
   # code for stats table
   def export_to_excel():
	   QtGui.QApplication.setOverrideCursor(QtCore.Qt.WaitCursor)
	   allrows = table3.rowCount()
	   allcolumns = table3.columnCount()
	   if allrows==0:
			msg.setText("Blank Table")
			msg.show()
			QtGui.QApplication.restoreOverrideCursor()
	   else:
		 try:
			npt = datetime.datetime.now()
			npt = re.sub('\s+|:|\.','-',str(npt))
			
			workbook = xlsxwriter.Workbook('IB_Pricing_Result'+str(npt)+'.xlsx')
			worksheet = workbook.add_worksheet()
			for row in range(0,1):
				for column in range(0,allcolumns):
					column_name = table3.horizontalHeaderItem(column).text()
					format = workbook.add_format({'bold': False, 'bg_color': 'white'})
					worksheet.write(row, column, str(column_name),format)
						
			for row in range(1,allrows+1):
				for column in range(0,allcolumns):
					try:
						col = table3.item(row-1, column).background().color().name()
						tx = table3.item(row-1,column).text()
						if col == "#008000":
							format = workbook.add_format({'bold': False, 'bg_color': 'green'})
							worksheet.write(row, column, str(tx),format)
						elif col == "#ff0000":
							format = workbook.add_format({'bold': False, 'bg_color': 'red'})
							worksheet.write(row, column, str(tx),format)
						elif col == "#ffff00":
							format = workbook.add_format({'bold': False, 'bg_color': 'yellow'})
							worksheet.write(row, column, str(tx),format)
						elif col == "#000000":
							format = workbook.add_format({'bold': False, 'bg_color': 'white'})
							worksheet.write(row, column, str(tx),format)
					
					except:
						tx = ''
						format = workbook.add_format({'bold': False, 'bg_color': 'white'})
						worksheet.write(row, column, str(tx),format)
					
					
						
					
			workbook.close()	
			msg.setText("Data exported to Excel")
			msg.show()
			QtGui.QApplication.restoreOverrideCursor()	
		 except:
			msg.setText("Error Occured")
			msg.show()
			QtGui.QApplication.restoreOverrideCursor()
	   
	   
	   
	   
	   
	   
	   
	   
   
   export_but1 = QPushButton("Export to excel",win4)
   export_but1.clicked.connect(export_to_excel)
   export_but1.setStyleSheet("background-color:lightgreen")
   export_but1.move(20,1)
   export_but1.setFixedWidth(120)
   
   
   
   def show_stats():
	   QtGui.QApplication.setOverrideCursor(QtCore.Qt.WaitCursor)
	   export_but1.show()
	   try:
		   db = pymysql.connect("35.154.31.255","crawl","St@nD@#d","scrap")
		   sql = """ select * from `Ib_Price`"""
		   stats_data = pd.read_sql(sql,db)
	   except:
		   	msg.setText("Error Occured")
			msg.show()
			QtGui.QApplication.restoreOverrideCursor()
	   
	   unique_sku = set(stats_data['IB Sku'])
	   unique_dates = set(stats_data['Date'])
	   headers = 'Values;'
	   for each in unique_dates:
		   headers = headers + each + ';'
	   
	   table3.setColumnCount(len(unique_dates)+1)
	   table3.setHorizontalHeaderLabels(QString(headers).split(";"))
	   table3.setColumnWidth(0,100)
	   wid = 1080/(len(unique_dates)+1)
	   for i in range (1,len(unique_dates)+1):
			table3.setColumnWidth(i,wid)
	   table3.setRowCount(len(unique_sku))
	   for i in range (len(unique_sku)):
			table3.setRowHeight(i,30)
	   for i,r in enumerate(unique_sku):
			table3.setItem(i,0,QTableWidgetItem(str(r)))
	   for row in stats_data.iterrows():
			date = row[1]['Date']
			sku = row[1]['IB Sku']
			
			if (str(row[1]["Moglix Price"])) == "nan":
				row[1]["Moglix Price"] = "-"
			if (str(row[1]["Amazon Price"])) == "nan":
				row[1]["Amazon Price"] = "-"
			if (str(row[1]["Snapdeal Price"])) == "nan":
				row[1]["Snapdeal Price"] = "-"
			
			if ((row[1]["IB Price"] < row[1]["Moglix Price"]) and (row[1]["IB Price"] < row[1]["Snapdeal Price"]) and (row[1]["IB Price"] < row[1]["Amazon Price"])):
						value = "green"
			else:
						value = "red"
			if ((row[1]["Moglix Price"] == "-") and (row[1]["Snapdeal Price"] == "-") and (row[1]["Amazon Price"] == "-")):
						value = "yellow"
						
			
			
			
			for j,k in enumerate(unique_sku):
				if k == sku:
					for l,m in enumerate(unique_dates):
						if m == date:
							table3.setItem(j,l+1,QTableWidgetItem(str(row[1]["IB Price"])))
							table3.item(j,l+1).setBackground(QtGui.QColor(value))
						else:
							pass
					break
				else:
					pass
			
			
			
			
			
			
			
	   table3.show()
	   win4.show()			
	   QtGui.QApplication.restoreOverrideCursor()
   
   
   sku_stats = QPushButton("SKU Stats",win)
   sku_stats.clicked.connect(show_stats)
   sku_stats.setStyleSheet("background-color:lightgreen")
   sku_stats.move(75,1)
   sku_stats.setFixedWidth(120)
   sku_stats.show()
   
   
   input_sku = QLineEdit(win)
   input_sku.setPlaceholderText("SKU")
   input_sku.setFixedWidth(150)
   input_sku.move(580,1)
   input_sku.show()
   
   def track():
		   QtGui.QApplication.setOverrideCursor(QtCore.Qt.WaitCursor)
		   try:
				db = pymysql.connect("35.154.31.255","crawl","St@nD@#d","scrap")
				sql = """select * from Ib_Price where `IB Sku` = """+"'"+str(input_sku.text())+"'"
				data = pd.read_sql(sql,db)
		   except:
			   msg.setText("Connection Problem")
			   msg.show()
			   QtGui.QApplication.restoreOverrideCursor()
		   
		   table1.setShowGrid(False)
		   table1.setRowCount(len(data))
		   for i in range (len(data)):
				table1.setRowHeight(i,50)
			
		   for index, row in data.iterrows():
				if (str(row["Moglix Price"])) == "nan" or (str(row["Moglix Price"])) == "None":
					row["Moglix Price"] = "-"
				if (str(row["Amazon Price"])) == "nan" or (str(row["Amazon Price"])) == "None":
					row["Amazon Price"] = "-"
				if (str(row["Snapdeal Price"])) == "nan" or (str(row["Snapdeal Price"])) == "None":
					row["Snapdeal Price"] = "-"
					
				table1.setItem(index,0,QTableWidgetItem(str(row["IB Sku"])))
				table1.setItem(index,1,QTableWidgetItem(str(row["IB Category"])))
				table1.setItem(index,2,QTableWidgetItem(str(row["IB Status"])))
				table1.setItem(index,3,QTableWidgetItem(str(row["IB Price"])))
				table1.setItem(index,4,QTableWidgetItem(str(row["Moglix Status"])))
				table1.setItem(index,5,QTableWidgetItem(str(row["Moglix Price"])))
				table1.setItem(index,6,QTableWidgetItem(str(row["Snapdeal Status"])))
				table1.setItem(index,7,QTableWidgetItem(str(row["Snapdeal Price"])))
				table1.setItem(index,8,QTableWidgetItem(str(row["Amazon Status"])))
				table1.setItem(index,9,QTableWidgetItem(str(row["Amazon Price"])))
				table1.setItem(index,10,QTableWidgetItem(str(row["Date"])))
				
				if ((row["IB Price"] < row["Moglix Price"]) and (row["IB Price"] < row["Snapdeal Price"]) and (row["IB Price"] < row["Amazon Price"])):
					for i in range(11):
						table1.item(index,i).setBackground(QtGui.QColor("lightgreen"))
				else:
					for i in range(11):
						table1.item(index,i).setBackground(QtGui.QColor("Red"))
				if ((row["Moglix Price"] == "-") and (row["Snapdeal Price"] == "-") and (row["Amazon Price"] == "-")):
					for i in range(11):
						table1.item(index,i).setBackground(QtGui.QColor("yellow"))
		   QtGui.QApplication.restoreOverrideCursor()
		   win1.show()
		   table1.show()
	   
   track_but = QPushButton("Track SKU",win)
   track_but.clicked.connect(track)
   track_but.setStyleSheet("background-color:POWDERBLUE")
   track_but.move(740,1)
   track_but.setFixedWidth(100)
   track_but.show()
  
   sys.exit(app.exec_())

if __name__ == '__main__':
   window()
