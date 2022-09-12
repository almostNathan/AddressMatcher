from distutils.command.build import build
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
import tkinter
from tkinter import filedialog
import os

def writeAddresses(sheet1, list1, sheet2, list2, matches):
    column1 = get_column_letter(sheet1.max_column)
    column2 = get_column_letter(sheet1.max_column)
    for item in matches:
        sheet1[column1+str(item[0])] = sheet2[column2+str(item[1])].value


#takes list if strings and returns 
def getAddresses(sheet, selectedCols):
    maxRows = sheet.max_row
    addressesFull = []
    zipList = []
    for rowNum in range(2,maxRows+1):
        addressesFull.append(sheet[selectedCols[0]+str(rowNum)].value)
        zipList.append(sheet[selectedCols[1]+str(rowNum)].value)
    addressesNum = list(map(lambda x: x.split()[0],addressesFull))

    returnAddressList=[]
    for i in range(0, maxRows-1):
        returnAddressList.append([addressesNum[i],zipList[i]])

    return returnAddressList
    

    

def buildColumnSelector(headerList):
    window = tkinter.Tk()
    window.title('Select Columns')
    window.geometry('500x500')

    label = tkinter.Label(window, bg='white', width=20, text='Select Address and Zip')
    label.pack()

    #initialize list to hold select options
    colList = []
    #list of tkinter.IntVar()'s
    for i in range(0,len(headerList)):
        colList.append(tkinter.IntVar())


    counter = 0
    for column in headerList:
        tkinter.Checkbutton(window, text=column, variable=colList[counter], onvalue= counter+1, offvalue=0).pack()
        counter+=1
        
    #reference
    #colA = tkinter.Checkbutton(window, text='A', variable=colList[0], onvalue= 1, offvalue=0).pack()

    btn = tkinter.Button(window, text='Submit', command = window.destroy).pack()


    window.mainloop()

    return colList

#takes a list of column numbers and returns list of letters
def getSelectedColLetters(headerList):
    #ask user what columns contain addresses
    selectedCols = buildColumnSelector(headerList)
    #remove not selected columns
    selectedCols = list(filter(lambda x: x.get()!=0,selectedCols))
    #get list of values as column LETTERS
    selectedCols = list(map(lambda x: get_column_letter(x.get()), selectedCols))

    return selectedCols

def getHeaders(sheet):
    headerList = []
    for col in range(1, sheet.max_column+1):
        cell = sheet[get_column_letter(col)+"1"].value
        headerList.append(cell)
    return headerList


#get the workbook filepath from the user, dialog fileselect window
workbookFilePath1 = "C:/Users/natha/Documents/BBB Address Matcher/AceTest.xlsx"
#workbookFilePath1 = filedialog.askopenfilename(initialdir=os.getcwd(), title="Select the Excel File", filetypes=(("Excel Files", "*.xlsx"),("All Files", "*.*")))

workbookFilePath2 = "C:/Users/natha/Documents/BBB Address Matcher/AceTest2.xlsx"
#workbookFilePath2 = filedialog.askopenfilename(initialdir=os.getcwd(), title="Select the Excel File", filetypes=(("Excel Files", "*.xlsx"),("All Files", "*.*")))


workbook1 = load_workbook(workbookFilePath1)
sheet1 = workbook1.active
headerList1 = getHeaders(sheet1)
selectedCols1 = getSelectedColLetters(headerList1)
addressZip1 = getAddresses(sheet1, selectedCols1)

workbook2 = load_workbook(workbookFilePath2)
sheet2 = workbook2.active
headerList2 = getHeaders(sheet2)
selectedCols2 = getSelectedColLetters(headerList2)
addressZip2 = getAddresses(sheet2, selectedCols2)

matches = []
for item1 in addressZip1:
    for item2 in addressZip2:
        if item1==item2:
           matches.append([addressZip1.index(item1),addressZip2.index(item2)]) 


writeAddresses(sheet1, addressZip1, sheet2, addressZip2, matches)


