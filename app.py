"""
this is the styling branch
"""

import pandas
import tkinter
from tkinter import Button
from tkinter import filedialog
from tkinter import Label
import xlsxwriter

"""
CONSTANTS:
Titles dictionaries: Columns names for feeder, and transformer sheets, program will look for these titles in the uploaded feeder and transformer excel sheets
"""
FEEDER_NAMES = {
        "FEEDER" : "اسم المغذي ورقمه",
        "STATION" : "اسم المحطة",
        "CITYSIDE" : "جانب المدينة",
        "TYPE" : "نوع المغذي",
        "LENGTH" : "SHAPE_Length",
        "STATUS" : "حالة المغذي",
        "NUMBER" : "رقم المغذي"
        }
TRANS_NAMES = {
        "FEEDER" : "اسم المغذي  ورقمه بالعربي",
        "SIZE" : "حجم المحولة",
        "TYPE" : "نوع المحولة",
        "STATUS" : "الحالة"
        }
CHECK_MARK = u'\u2713' # This constant is used by the tk interface for displaying check mark sign for user messages

""" 
Global variables:
These variables are updated in the functions, but they must be declared as global at the begining of functions
"""
feedFrame = None # contains the frame read from feeders excel
transFrame = None #contains the frame read from transformers excel
userMessage = None # contains message displayed 
feederFlag = False # will be true if the the feeder file uploaded and processed successfully
transFlag = False # will be true if the the transformers file uploaded and processed successfully

"""
Each station11k will contains its info and multiple unique feeder objects
"""
class Station11K:
    stationsDic = {}
    def __init__(self, name, citySide):
        self.name = name
        self.citySide = citySide
        self.feedersList = []
        Station11K.stationsDic[name] = self
    def addFeeder(self, feeder):
        self.feedersList.append(feeder)

"""
Each feeder object contains all info related to the feeder
"""
class Feeder:
    objectsDic = {} # collection of current objects in the class for search
    def __init__(self, name):
        self.name = name
        self.cableLength = 0
        self.overLength = 0
        self.citySide = ""
        self.number = ""
        self.trans = {
                "kiosk": {"100":0, "250":0, "400":0, "630":0, "1000":0, "other":0},
                "indoor": {"100":0, "250":0, "400":0, "630":0, "1000":0, "other":0},
                "outdoor": {"100":0, "250":0, "400":0, "630":0, "1000":0, "other":0}
                }
        Feeder.objectsDic[name] = self
    def totalLength(self):
        return self.cableLength + self.overLength

"""
Functionality: Compare the dsired column headers (stored in the program) with the column headers read from excel sheet
Parameters:
    (1) dictionary: the desired columns headers (stored in the program)
    (2) list: headers read from the excel file
return: boolean
    True: If stored headers, and headers read from the excel file are the same
    False: If stored headers, and headers read from the excel file are not the same
"""
def validate_columns(desiredHeaders, fileHeaders):
    for key, value in desiredHeaders.items():
        if value not in fileHeaders:
            return False
    return True

""""
Functionality: 
    Read data from feeder excel file, put the data in both Stations11k, and Feeder classes.
    only the data of feeders which there status is 'بالعمل' are read, others will be neglected
Return:
    nothing!
"""
def import_feeders():
    """ include the global variables related to feeder, if not included, they can't be modified inside the function """
    global feedFrame
    global feederFlag
    """ Create constant variables instead of using the dictionary, make it cleaner and also easier to maintain in the future. """
    FEEDER_FEEDER = FEEDER_NAMES["FEEDER"]
    FEEDER_STATION = FEEDER_NAMES["STATION"]
    FEEDER_CITYSIDE = FEEDER_NAMES["CITYSIDE"]
    FEEDER_TYPE = FEEDER_NAMES["TYPE"]
    FEEDER_LENGTH = FEEDER_NAMES["LENGTH"]
    FEEDER_STATUS = FEEDER_NAMES["STATUS"]
    FEEDER_NUMBER = FEEDER_NAMES["NUMBER"]
    """ Upload excel file contain the feeders data """
    try:
        filename = filedialog.askopenfilename(initialdir = "/",title = "اختر ملف المغذيات",filetypes = (("Excel files","*.xls"),("all files","*.*")))
    except:
        userMessage.configure(text="خطأ اثناء تحميل ملف المغذيات", fg="red")
        feederFlag = False
        return
    feedFrame = pandas.read_excel(filename,sheet_name=0) # Create panda fram  reading excel file
    columnsHeaders = feedFrame.columns.tolist()  # Create a list contains all column header of the excel sheet
    """ Validate the headers of the excel sheet """
    if not validate_columns(FEEDER_NAMES,columnsHeaders):
        userMessage.configure(text="هنالك عدم مطابقة في عناوين الاعمدة في ملف المغذيات", fg="red")
        feederFlag = False
        return
    """ 
    Read the excel sheet (stored in pandas frame) row by row, and store result in Station11k, and Feeder classes
    rows will be neglected if the status is not (بالعمل) 
    """
    try:
        for index, row in feedFrame.iterrows():
            if row[FEEDER_STATUS] == "بالعمل":
                feederName = str(row[FEEDER_FEEDER]).strip() # remove leading spaces from the feeder name
                """ 
                check if the feeder was previously read, 
                if yes, then the data will be read and stored in the same feeder, 
                if not, a new feeder will be created, and then data stored in it 
                """
                feeder = Feeder.objectsDic.get(feederName, None)
                if feeder is None: # If feeder is not previously read from another row in the sheet
                    feeder = Feeder(feederName)
                    feeder.number = row[FEEDER_NUMBER]
                    citySide = row[FEEDER_CITYSIDE]
                    feeder.citySide = citySide
                    stationName = row[FEEDER_STATION]
                    station = Station11K.stationsDic.get(stationName, None)
                    if station is None:
                        station = Station11K(stationName, citySide)
                    station.addFeeder(feeder)
                if row[FEEDER_TYPE] == "overhead":
                    feeder.overLength = round(row[FEEDER_LENGTH],2) # keep only two digits after the dot
                elif row[FEEDER_TYPE] == "cable":
                    feeder.cableLength = round(row[FEEDER_LENGTH],2) # keep only two digits after the dot
                else:
                    print(f"Feeder {row[FEEDER_FEEDER]} has ({row[FEEDER_TYPE]}) type field")
        userMessage.configure(text=f"تمت معالجة ملف المغذيات {CHECK_MARK}", fg="green") # Display success message to user
        feederFlag = True  # data can be processed by the feeder processing functions
    except:
        userMessage.configure(text="حدث خطأ اثناء معالجة بيانات ملف المغذيات", fg="red") # Display failure message to user
        feederFlag = False # data will not be processed by the feeder processing functions

""""
Functionality: 
    Read data from transformers excel file row by row, looks for transformers data of each existing feeder (in Feeder class)
    only the data of transformers with status 'good' will be added, others will be neglected
Return:
    nothing!
"""
def import_transformers():
    """ include the global variables related to feeder, if not included, they can't be modified inside the function """
    global transFrame
    global transFlag
    """ Create constant variables instead of using the dictionary, make it cleaner and also easier to maintain in the future. """
    TRANS_FEEDER = TRANS_NAMES["FEEDER"]
    TRANS_SIZE = TRANS_NAMES["SIZE"]
    TRANS_TYPE = TRANS_NAMES["TYPE"]
    TRANS_STATUS = TRANS_NAMES["STATUS"]
    """ Upload excel file contain the feeders data """
    try:
        filename = filedialog.askopenfilename(initialdir = "/",title = "اختر ملف المحولات",filetypes = (("Excel files","*.xls"),("all files","*.*")))
    except:
        userMessage.configure(text="خطأ اثناء تحميل ملف المحولات", fg="red")
        transFlag = False
        return
    transFrame = pandas.read_excel(filename,sheet_name=0) # Create panda fram  reading excel file
    headers = transFrame.columns.tolist() # Create a list contains all column header of the excel sheet
    """ Validate the headers of the excel sheet """
    if not validate_columns(TRANS_NAMES,headers):
        userMessage.configure(text="هنالك عدم مطابقة في عناوين الاعمدة في ملف المحولات", fg="red")
        transFlag = False
        return
    """ 
    Read the excel sheet (stored in pandas frame) row by row, and store result in feeders
    rows will be neglected if the status is not (good) 
    """
    try:        
        for index, row in transFrame.iterrows():
            if row[TRANS_STATUS] == "good":
                name = str(row[TRANS_FEEDER]).strip() # remove leading spaces from the feeder name
                feeder = Feeder.objectsDic.get(name, None) # check if the feeder already exist in the feeders list
                """ if feeder exist, add transformers data to it, if not, ignore it. """
                if feeder is not None:
                    transType = row[TRANS_TYPE]
                    transSize = row[TRANS_SIZE]
                    """ 
                    if trans has type and size, add it to its proper place trans[type][size]
                    if trans only has type, the add it to trans[type]['other']
                    if trans doesn't has type, then ignore it.
                    """
                    if transSize in ['100','250','400','630','1000'] and transType in ['indoor', 'outdoor', 'kiosk']:
                        feeder.trans[transType][transSize] += 1
                    else:
                        if transType in ['indoor', 'outdoor', 'kiosk']:
                            feeder.trans[transType]['other'] += 1
        userMessage.configure(text=f"تمت معالجة ملف المحولات {CHECK_MARK}", fg="green") # user success message
        transFlag = True # data can be processed by the feeder processing functions
    except:
        userMessage.configure(text="حدث خطأ اثناء معالجة بيانات ملف المحولات", fg="red") # user failure message
        transFlag = False # data will not be processed by the feeder processing functions
     
"""
Functionality:
    if both excel files of feeders, and transformers are uploaded, and processed properly, 
    this method will create an excel file, and puts in it all the data required for the ministery report.
Returns:
    Nothing !
"""
def export_ministery_report():
    """
    First check whether the two excel files were uploaded and processed properly,
    if not, the method will stop and ask user to upload and process the proper files 
    """
    if not feederFlag and not transFlag:
        userMessage.configure(text="تأكد من تحميل الملفات بصورة صحيحة قبل محاولة تصديرها", fg="red")
        return
    try:
        """ get a file name from user browsing box """
        filename = filedialog.asksaveasfilename(filetypes=(("Excel files", "*.xlsx"),("All files", "*.*") ))
        if filename is None: # asksaveasfile return `None` if dialog closed with "cancel".
            return
        """ create excel file workbook, and a worksheet, and customize the worksheet """
        workbook = xlsxwriter.Workbook(filename + ".xlsx")
        worksheet = workbook.add_worksheet()
        worksheet.right_to_left() # make it arabic oriented
        worksheet.set_zoom(70) # the zoom will be 70%
        """ 
        Create cell style per each type of cells
        style (format) will be added per cell, because I found it eaiser to perform in XlsxWriter
        """
        titleCellFormat = workbook.add_format({'bold': True, 'font_size':14, 'align': 'center', 'valign':'vcenter', 'border':True, 'pattern':1, 'bg_color':'#d3d3d3'})
        logoCellFormat = workbook.add_format({'bold': True, 'font_size':18, 'align': 'center', 'valign':'vcenter', 'border':True})
        genericCellFormat = workbook.add_format({'align': 'center', 'valign':'vcenter', 'border':True})
        sumCellFormat = workbook.add_format({'bold': True, 'font_size':14, 'align': 'center', 'valign':'vcenter', 'border':True})
        """ set the width of columns """
        worksheet.set_column("A:A",18)
        worksheet.set_column("B:B",12)        
        worksheet.set_column("C:C",20)
        worksheet.set_column("D:F",12)
        worksheet.set_column("Y:Y",15)
        """
        Build title and log, which will be first 4 rows
        """
        """ 1st row for logo image """
        worksheet.merge_range("A1:Y1", "", genericCellFormat) 
        worksheet.set_row(0,210)
        worksheet.insert_image('A1', 'logo.png', {'x_scale': 1.451, 'y_scale': 1.451})
        """ 2nd row for department title """
        worksheet.merge_range("A2:Y2", "مديرية توزيع كهرباء مركز نينوى", logoCellFormat)
        worksheet.set_row(1,40)
        """ 3rd and 4th rows for columns titles """
        worksheet.set_row(2,25)
        worksheet.set_row(3,25)
        for cellRange, text in (["A3:A4","اسم المحطة"],["B3:B4","جانب المدينة"],["C3:C4","اسم المغذي"]):
            worksheet.merge_range(cellRange, text, titleCellFormat)
        worksheet.merge_range("D3:F3", "اطوال المغذيات - بالمتر", titleCellFormat)
        for cellRange, text in (["D4","ارضي"],["E4","هوائي"],["F4","الكلي"]):
            worksheet.write(cellRange, text, titleCellFormat)
        titleColumnIndex = 6 # transformers' columns titles start at the 7th column (first 6 columns taken for station, city side, feeder length, etc.)
        for cellRange, text in (["G3:L3","محولات صندوقية"],["M3:R3","غرف محولات"],["S3:X3","محولات هوائية"]):
            worksheet.merge_range(cellRange, text, titleCellFormat)
            for size in [100, 250, 400, 630, 1000, "اخرى"]:
                worksheet.write(3, titleColumnIndex, size, titleCellFormat)
                titleColumnIndex += 1
        worksheet.merge_range("Y3:Y4","مجموع المحولات", titleCellFormat)
        """ 
        build data cells, loop through station, and put its feeders as rows.
        we need two pointers, a pointer to the first row in the station, and point to the end row of the station, 
        they will be used to (1) put the station name, and city side in merged cells their height equal to the number of rows in that station
        and (2) know where the second station starts.
        """
        startRowIndex = 4 # starts at row number 5, because first 4 rows were taken for image, and titles
        endRowIndex = 4 # before each new station, start and end pointers should be pointer at same row
        """ 
        initiate variables used to sum the data needed at the end of the sheet 
        these will be updated when looping through stations and feeders
        """
        totalFeeders = 0
        totalCableLength = 0
        totalOverLength = 0
        totalCombinedLength = 0
        totalTransTypes = {
                "kiosk": {"100":0, "250":0, "400":0, "630":0, "1000":0, "other":0},
                "indoor": {"100":0, "250":0, "400":0, "630":0, "1000":0, "other":0},
                "outdoor": {"100":0, "250":0, "400":0, "630":0, "1000":0, "other":0}
                }
        grandTransSum = 0 # this is the summation of all transformers in all feeders
        """ loop through feeders for each station """
        for name, station in Station11K.stationsDic.items():
            feedersList = station.feedersList # can't use the sort function on the list if I don't store it in a variable first
            feedersList.sort(key=lambda x: x.number, reverse=False) # sorting feeders inside a station according to their numbers
            columnIndex = 2 # first two columns are taken for station name, and city side
            for feeder in feedersList:
                for text in (feeder.name, feeder.cableLength, feeder.overLength, feeder.totalLength()):
                    worksheet.write(endRowIndex, columnIndex, text, genericCellFormat)
                    columnIndex += 1
                sumTransRow = 0 # sum the total transformers (all types) in each feeder (i.e. stations in each row)
                colors = {'kiosk':'#fef200', 'indoor':'#75d86a', 'outdoor':'#4dc3ea'} # coloring each type of stations
                for shape in ['kiosk', 'indoor', 'outdoor']:
                    color = colors[shape] # when put colors[shape] in the format directly, I faced error in the program, so I put it first in variable then, used it
                    """ I can't put below fromat at the begining of functions as other formats, because it useds the color variable, which is generated inside this loop"""
                    transCellFormat = workbook.add_format({'align': 'center', 'valign':'vcenter', 'border':True, 'pattern':1, 'bg_color':color}) 
                    for size in ['100', '250', '400', '630', '1000', 'other']:
                        sumTransType = feeder.trans[shape][size] # the sumation of each type of transformer in one feeder
                        worksheet.write(endRowIndex, columnIndex, sumTransType, transCellFormat)
                        columnIndex += 1
                        sumTransRow += sumTransType # add the transformers of specific type to the sumation of transformers in the current feeder
                        totalTransTypes[shape][size] += sumTransType # add the transformers of specific type to the total transformers of this type (in all feeders)
                worksheet.write(endRowIndex, columnIndex, sumTransRow, genericCellFormat)
                """ update the total variables """
                totalFeeders += 1
                totalCableLength += feeder.cableLength
                totalOverLength += feeder.overLength
                columnIndex = 2 # reset column index for each new feeder
                endRowIndex += 1 # end row index refer to next empty row
            worksheet.merge_range(startRowIndex,0,endRowIndex-1,0, name, genericCellFormat) # add the station in the first column, with height equal all feeder rows
            worksheet.merge_range(startRowIndex,1,endRowIndex-1,1, station.citySide, genericCellFormat) # add the city side in the first column, with height equal all feeder rows           
            worksheet.merge_range(endRowIndex,0,endRowIndex,24, "", titleCellFormat) # create an empty row, works as separation between stations
            endRowIndex += 1 # increase the row pointer to point to the next row after the empty one added.
            startRowIndex = endRowIndex # At the end of each new loop, the row start and end indexes should be equal
        """ finally, add the sumation row at the bottom of the sheet """
        columnIndex = 0
        totalCombinedLength = totalCableLength + totalOverLength
        for text in ["المجموع الكلي", "", totalFeeders, totalCableLength, totalOverLength, totalCombinedLength]:
            worksheet.write(endRowIndex, columnIndex, text, sumCellFormat)
            columnIndex += 1
        for shape in ['kiosk', 'indoor', 'outdoor']:
            for size in ['100', '250', '400', '630', '1000', 'other']:
                totalTransCol = totalTransTypes[shape][size]
                worksheet.write(endRowIndex, columnIndex, totalTransCol, sumCellFormat)
                grandTransSum += totalTransCol
                columnIndex += 1
        worksheet.write(endRowIndex, 24, grandTransSum, sumCellFormat) 
        worksheet.set_row(endRowIndex, 40) # set the height of row, I couldn't do at the beginning with other formats becuase it uses a variable the its value couldn't be known at the beginning
        workbook.close() # finally save the excel file
    except:
        userMessage.configure(text="حدث خطأ اثناء تصدير الملف", fg="red") # user message if any thing went wrong during executing the function

def export_transformers_text():
    if not feederFlag and not transFlag:
        userMessage.configure(text="تأكد من تحميل الملفات بصورة صحيحة قبل محاولة تصديرها", fg="red")
        return
    feederName, feederType, transformer  = [],[],[]
    for name, feeder in Feeder.objectsDic.items():
        feederName.append(feeder.name)
        feederType.append("cable")
        id_100 = feeder.trans['indoor']['100'] + feeder.trans['kiosk']['100']
        id_250 = feeder.trans['indoor']['250'] + feeder.trans['kiosk']['250']
        id_400 = feeder.trans['indoor']['400'] + feeder.trans['kiosk']['400']
        id_630 = feeder.trans['indoor']['630'] + feeder.trans['kiosk']['630']
        id_1000 = feeder.trans['indoor']['1000'] + feeder.trans['kiosk']['1000']
        id_other = feeder.trans['indoor']['other'] + feeder.trans['kiosk']['other']
        idStrig = f"100x{id_100}+250x{id_250}+400x{id_400}+630x{id_630}+1000x{id_1000}+اخرىx{id_other}"
        transformer.append(idStrig)            
        od_100 = feeder.trans['outdoor']['100']
        od_250 = feeder.trans['outdoor']['250']
        od_400 = feeder.trans['outdoor']['400']
        od_630 = feeder.trans['outdoor']['630']
        od_1000 = feeder.trans['outdoor']['1000']
        od_other = feeder.trans['outdoor']['other']
        od_total = od_100 + od_250 + od_400 + od_630 + od_1000 + od_other
        if od_total>0:
            feederName.append(name)
            feederType.append("overhead")
            odString = f"100x{od_100}+250x{od_250}+400x{od_400}+630x{od_630}+1000x{od_1000}+اخرىx{od_other}"
            transformer.append(odString)
    return {"Feeder":feederName, "Type":feederType, "transformer":transformer}

def main():
    global userMessage
    window = tkinter.Tk()
    window.title("تقارير قسم المعلوماتية V1")
    window.geometry("400x380")
    space1 = Label(window, text="")
    space1.pack()
    feederButton = Button(window, text="تحميل جدول المغذيات (ملف أكسل)",padx='75', pady='15', command=import_feeders)
    feederButton.pack()
    space2 = Label(window, text="")
    space2.pack()
    transButton = Button(window, text="تحميل جدول المحولات (ملف أكسل)",padx='75', pady='15', command=import_transformers)
    transButton.pack()
    space3 = Label(window, text="")
    space3.pack()    
    exportMinistry = Button(window, text="تصدير تقرير الوزارة",padx='15', pady='15', command=export_ministery_report)
    exportMinistry.pack()
    space4 = Label(window, text="")
    space4.pack()    
    exportTrans = Button(window, text="تصدير عدد المحولات",padx='15', pady='15', command=export_transformers_text)
    exportTrans.pack()
    space5 = Label(window, text="")
    space5.pack()
    userMessage = Label(window, text="", fg="red", font=("Helvetica", 12))
    userMessage.pack()
    window.mainloop()

if __name__ == '__main__':
    main()
