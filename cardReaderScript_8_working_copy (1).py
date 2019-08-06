
# Impoting xrld, date and openpyxl modules.

import xlrd
from xlutils.copy import copy
#import openpyxl   #probably can delete
import datetime
import os
global wb
global wbfinal
global finalcolumn
global finalrow

##################################################################################
#Initalizing all the necessary variables, defining the file paths and creating a new day's .txt file.
#STEP 1
tracker = 12 #MUST CHANGE BACK TO 27
qcode = 0
name = "unknown"
finalcolumn = 0
finalrow = 0

##################################################################################
datenow = datetime.datetime.now()
#print(datenow)
daynow = str(datenow.day)      #Want to make this MMM format not number %b
monthnow = str(datenow.month)
dateofclass0 = datenow.strftime("%d/%m/%Y")
#print(dateofclass)
filedate = daynow + monthnow
#print(filedate)

##################################################################################
#Need to make sure you create a folder in C:\ named Project_Test_Folder
#Within Project_Test_Folder place 'classDecoder.xlsx' and 'classAttendance.xlsx'

textformat = 'txt'
folderpath = os.path.abspath('C:\Project_Test_Folder')
filename = os.path.join(folderpath,filedate + "." + textformat)
loc = open(filename, "a+")
decoder = "classDecoder.xlsx"
decoderloc = os.path.join(folderpath,decoder)
attendance = "classAttendance1.xls"
attendanceloc = os.path.join(folderpath,attendance)
wb = xlrd.open_workbook(decoderloc)
wbfinal = xlrd.open_workbook(attendanceloc)
##################################################################################
#markAttendace is a funtion that searches the classAttendance.csv for both the name and the date obtained and marks "X" in the intersecting cell annotating the students attendance.
#STEP 4
def markAttendance(name2):
  wbfinal = xlrd.open_workbook(attendanceloc)
  sheetfinal = wbfinal.sheet_by_index(0)
  finalcolumn = 0
  finalrow = 0
  for row2 in range(sheetfinal.nrows):
          #print(sheetfinal.nrows)
          #print(row2)
          for col2 in range(sheetfinal.ncols):
                  #print(sheetfinal.ncols)
                  #print(col2)
                  myCellfinal = sheetfinal.cell(row2, col2)
                  myCellfinal2 = str(myCellfinal)
                  dateofclass = str(dateofclass0)
                  #mycellfinal3 = xlrd.xldate.xldate_as_datetime(myCellfinal[0].valu, book.datemode)
                  #print(myCellfinal2)
                  #print(myCellfinal3)
                  #print(name2)
                  #print(dateofclass)
                  #for q in range(sheet.ncols):
                  if name2 == myCellfinal2:
                          print(name2)
                          finalrow = row2
                          #print(finalrow)
        #nfrow = sheet.cell_column(q)
#for q in range(sheet.ncols):
                  elif dateofclass in myCellfinal2:
                          finalcolumn = col2
                          print(dateofclass)
                          #print(finalcolumn)
                  #row2 = row2 + 1
                  #col2 = col2 + 1
                        #else: return()
        #daterow = sheet.cell_column(q)
  #print(finalcolumn)
  #print(finalrow)
  #if finalcolumn or finalrow == 0: return()
  #else:
  #wbfinal.close()
  writewbfinal= copy(wbfinal)
  sheet1 = writewbfinal.get_sheet(0)
  print(finalrow, ",", finalcolumn)
  sheet1.write(finalrow,finalcolumn, 'X')
  writewbfinal.save("C:\Project_Test_Folder\classAttendance1.xls")
  #c1 = sheetfinal.cell(finalrow, finalcolumn)
  #c1.value = "X"
  return()

###################################################################################
#GetName is a function that compares the input (renamed as qlcode) to the decoder csv file in order to get an output of the rfid card's owner's name.
#STEP 3
def getName(qlcode):
  myCell2 = " "
  wb = xlrd.open_workbook(decoderloc)
#Made wb Global
  sheet = wb.sheet_by_index(0)
  #sheet.cell_value(0,0)
  for row in range(sheet.nrows):
    for col in range(sheet.ncols):
                        #print(row)
                        qlcode = str(qlcode)  #no idea why but the code breaks if i dont turn qlcode into a string again.
                        myCell = sheet.cell(row, col)
                        myCell2 = str(myCell)
                        #print(qlcode)
                        #print(myCell2)
                        #for i in range(sheet.ncols):
                        if qlcode in myCell2:
                                name1 = sheet.cell(row, col + 1)
                                name = str(name1)
                                print(myCell2)
                                #nrow = sheet.cell_row(i)
                                #name = sheet.cell(nrow,1)
                                print(name)
        ####add last function here and return nothing!!!!
                                markAttendance(name)
  return ()
        
      #else: break
##################################################################################
#Records the 27 student class's qcodes and timestamps on a text file to be used in the GetName Function next.
#STEP 2
while tracker > 0:
  current = datetime.datetime.now()
  currenthour = current.hour
  if currenthour == 9: break
  else:
    qcod = input("Please Place Your SMU Card on the Reader")
    qcode = str(qcod)
    if qcode == False: break
    getName(qcode)
    datenow = datetime.datetime.now()
    #datenowFormat = datenow.strftime("%d/%m/%Y")
    qdate = qcode + "\n" + str(datenow) + "\n"
    #print(datenowFormat)
    if qcode not in loc: #pretty sure this is always going to be a no, path not in path
      loc.write(qdate)
      tracker -= 1

##################################################################################
# Running the getName Function withing the .txt that was written today for class.
#STEP 3


##################################################################################
#STEP 5


loc.close()
#wb.close()
#wbfinal.close()
