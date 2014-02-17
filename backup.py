import xlrd, xlwt
import datetime, time
from time import gmtime, strftime
import os, sys
from xlutils.copy import copy

target = ""
result = ""

av = sys.argv
own = "OK"
if len(av) > 1:
	own = av[1].upper()

class Backup:
	def __init__(self, path_excel, own):
		self.path_excel = path_excel
		self.own = own
		self.create_paths_folders()
		
	def create_paths_folders(self):
		try:
			workbook = xlrd.open_workbook(self.path_excel)
			Wtworkbook = copy(workbook)
			today = datetime.datetime.now()
			worksheet = workbook.sheet_by_index(int(today.strftime('%m'))-1)
			Wtworksheet  = Wtworkbook.get_sheet(int(today.strftime('%m'))-1)
			num_rows = worksheet.nrows
			for row in range(3,num_rows):
				file = worksheet.cell_value(row,1).replace("\\","/")+"/"+worksheet.cell_value(row,3)
				if os.path.isfile(file):
					if strftime("%Y%m%d", gmtime(os.path.getmtime(file))) == strftime("%Y%m%d", gmtime()):
						Wtworksheet.write(row,3+int(strftime("%d", gmtime())),self.own)
					else:
						Wtworksheet.write(row,3+int(strftime("%d", gmtime())),"OLD")
						print "Fail: old file  "+file
				else:
					Wtworksheet.write(row,3+int(strftime("%d", gmtime())),"NO EXIST")
					print "Fail: file doesn't match  "+file
			workbook.release_resources()
			Wtworkbook.save(result)
		except Exception, e:
			print e
				
		
print "Arranca la verificacion del Backup"
Backup(target,own)
