import pandas
import xlsxwriter
from datetime import datetime
import tkinter
from tkinter import Frame, Button, PhotoImage, Label, LabelFrame, LEFT, RIGHT, NE, filedialog

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
SOURCE_NAMES = {
        "NAME" : "اسم المغذي ورقمه",
        "STATION_33" : "sub_name33",
        "STATION_132" : "اسم المحطة132 التي تغذيها",
        "STATUS" : "الحالة",
		"OPERATION": "حالة المغذي",
		"LENGTH" : "SHAPE.STLength()",
        "NUMBER" : "رقم المغذي",
        "CITYSIDE": "جانب المدينة"
		}
LOAD_NAMES = {
        "LOAD" : "AMPs",
        "VOLTS" : "الفولطية",
        "FEEDER" : "اسم المغذي ورقمه",
        }

CHECK_MARK = u'\u2713' # This constant is used by the tk interface for displaying check mark sign for user messages

""" 
Global variables:
These variables are updated in the functions, but they must be declared as global at the begining of functions
"""
userMessage = None # contains message displayed
loadMessage11K = None # message indicates if feeder (11 KV) is loaded or not
loadMessage33K = None # message indicates if source (33 KV) is loaded or not
feederFlag = False # will be true if the the feeder file uploaded and processed successfully
transFlag = False # will be true if the the transformers file uploaded and processed successfully
loadFlag = False # will be true if the the loads file uploaded and processed successfully
sourceFlag = False # will be true if the the sources file uploaded and processed successfully

"""
Each station11k will contains its info and multiple unique feeder objects
"""
class Station11K:
    stationsDic = {}
    def __init__(self, name, citySide):
        self.name = name
        self.citySide = citySide
        self.feedersList = []
        self.sourcesList = []
        Station11K.stationsDic[name] = self
    def addFeeder(self, feeder):
        self.feedersList.append(feeder)
    def addSource(self, source):
        self.sourcesList.append(source)

"""
Each source object contains all info related to the source
"""
class Source:
    objectsDic = {} # collection of current objects in the class to search if the source already exist before adding it
    def __init__(self, name):
        self.name = name
        self.station132 = ""
        self.length = 0
        self.number = ""
        self.load = 0
        self.volts = "33 KV"
        Source.objectsDic[name] = self # add the new source to the collection of the class

"""
Each feeder object contains all info related to the feeder
"""
class Feeder:
    objectsDic = {} # collection of current objects in the class to search if the feeder already exist before adding it
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
        self.load = 0
        self.volts = "11 KV"
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
    global feederFlag
    global loadMessage11K
    global userMessage
    """ if current date exceeded the expiry date, the program will show error message and stops working """
    if not validate_date():
        userMessage.configure(text="هنالك خطأ في البرنامج, اتصل بالمصصم على الرقم 07701791983 ", fg="red")
        return
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
        feedFrame = pandas.read_excel(filename,sheet_name=0) # Create panda fram  reading excel file
    except:
        feederFlag = False
        userMessage.configure(text="لم يتم اختيار ملف المغذيات", fg="red") # Display failure message to user
        loadMessage11K.configure(text="X", fg="red")
        return
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
        loadMessage11K.configure(text=f"{CHECK_MARK}", fg="green")
        userMessage.configure(text=f"تمت معالجة ملف المغذيات ", fg="green") # Display success message to user
        feederFlag = True  # data can be processed by the feeder processing functions
    except:
        loadMessage11K.configure(text="X", fg="red")
        userMessage.configure(text="حدث خطأ اثناء معالجة بيانات ملف المغذيات", fg="red") # Display failure message to user
        feederFlag = False # data will not be processed by the feeder processing functions
    return

"""
Functionality: 
    Read data from transformers excel file row by row, looks for transformers data of each existing feeder (in Feeder class)
    only the data of transformers with status 'good' will be added, others will be neglected
Return:
    nothing!
"""
def import_transformers():
    """ include the global variables related to transformers, if not included, they can't be modified inside the function """
    global transFlag
    global loadMessageTrans
    global userMessage
    """ if current date exceeded the expiry date, the program will show error message and stops working """
    if not validate_date():
        userMessage.configure(text="هنالك خطأ في البرنامج, اتصل بالمصصم على الرقم 07701791983 ", fg="red")
        return
    """ transformers file depends on feeders 11 KV file, so the feeder file must be uploaded first"""
    if not feederFlag:
        userMessage.configure(text="قم بتحميل جدول مغذيات (11 كف) اولا  ", fg="red")
        return
    """ Create constant variables instead of using the dictionary, make it cleaner and also easier to maintain in the future. """
    TRANS_FEEDER = TRANS_NAMES["FEEDER"]
    TRANS_SIZE = TRANS_NAMES["SIZE"]
    TRANS_TYPE = TRANS_NAMES["TYPE"]
    TRANS_STATUS = TRANS_NAMES["STATUS"]
    """ Upload excel file contain the feeders data """
    try:
        filename = filedialog.askopenfilename(initialdir = "/",title = "اختر ملف المحولات",filetypes = (("Excel files","*.xls"),("all files","*.*")))
        transFrame = pandas.read_excel(filename,sheet_name=0) # Create panda fram  reading excel file
    except:
        userMessage.configure(text="لم يتم تحميل ملف المحولات", fg="red")
        transFlag = False
        loadMessageTrans.configure(text="X", fg="red")
        return
    headers = transFrame.columns.tolist() # Create a list contains all column header of the excel sheet
    """ Validate the headers of the excel sheet """
    if not validate_columns(TRANS_NAMES,headers):
        userMessage.configure(text="هنالك عدم مطابقة في عناوين ملف المحولات", fg="red")
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
        loadMessageTrans.configure(text=f"{CHECK_MARK}", fg="green")
        userMessage.configure(text=f"تمت معالجة ملف المحولات", fg="green") # user success message
        transFlag = True # data can be processed by the feeder processing functions
    except:
        loadMessageTrans.configure(text="X", fg="red")
        userMessage.configure(text="حدث خطأ اثناء معالجة ملف المحولات", fg="red") # user failure message
        transFlag = False # data will not be processed by the feeder processing functions
    return

""""
Functionality: 
    Read data from sources (33 KV) excel file row by row, looks for sources (33 KV) data and stores the result in Source Class  
    only the data of sources with status 'good' and 'بالعمل' will be added, others will be neglected
Return:
    nothing!
"""
def import_sources():
    """ include the global variables related to transformers, if not included, they can't be modified inside the function """
    global sourceFlag
    global loadMessage33K
    global userMessage
    """ if current date exceeded the expiry date, the program will show error message and stops working """
    if not validate_date():
        userMessage.configure(text="هنالك خطأ في البرنامج, اتصل بالمصصم على الرقم 07701791983 ", fg="red")
        return
    """ Create constant variables instead of using the dictionary, make it cleaner and also easier to maintain in the future. """    
    NAME = SOURCE_NAMES["NAME"]
    STATION_33 = SOURCE_NAMES["STATION_33"]
    STATION_132 = SOURCE_NAMES["STATION_132"]
    STATUS = SOURCE_NAMES["STATUS"]
    OPERATION = SOURCE_NAMES["OPERATION"]
    LENGTH = SOURCE_NAMES["LENGTH"]
    NUMBER = SOURCE_NAMES["NUMBER"]
    CITYSIDE = SOURCE_NAMES["CITYSIDE"]
    try:
        filename = filedialog.askopenfilename(initialdir = "/",title = "اختر ملف المصادر",filetypes = (("Excel files","*.xls"),("all files","*.*")))
        sourceFrame = pandas.read_excel(filename,sheet_name=0) # create panda frame by reading excel file
    except:
        userMessage.configure(text="لم يتم تحميل ملف المصادر", fg="red")
        sourceFlag = False
        loadMessage33K.configure(text="X", fg="red")
        return
    headers = sourceFrame.columns.tolist() # Create a list contains all column header of the excel sheet
    """ Validate the headers of the excel sheet """
    if not validate_columns(SOURCE_NAMES,headers):
        userMessage.configure(text="هنالك عدم مطابقة في عناوين ملف المصادر", fg="red")
        transFlag = False
        return
    """ 
    Read the excel sheet (stored in pandas frame) row by row, and store result in Source class objects
    rows will be neglected if the status is not (good) or not operational 
    """
    try:
        for index, row in sourceFrame.iterrows():
            if row[STATUS] == "good" and row[OPERATION] == "بالعمل":
                sourceName = str(row[NAME]).strip() # remove leading spaces from the source name
                """ 
                check if the source was previously read, 
                if yes, then the data will be read and stored in the same source, 
                if not, a new feeder will be created, and then data stored in it 
                """
                source = Source.objectsDic.get(sourceName, None)
                if source is None: # If source is not previously read from another row in the sheet
                    source = Source(sourceName)
                    source.number = row[NUMBER]
                    citySide = row[CITYSIDE]
                    station33Name = row[STATION_33]
                    station = Station11K.stationsDic.get(station33Name, None)
                    """ 
                    if Station33Kv is not already exist, then ignore it, 
                    This to avoid a problem when station exists without having 11KV feeders
                    """
                    if station is not None:
                        station.addSource(source)
                    source.station132 = row[STATION_132]
                    source.length = round(row[LENGTH],2)
        loadMessage33K.configure(text=f"{CHECK_MARK}", fg="green")
        userMessage.configure(text=f"تمت معالجة ملف المصادر ", fg="green") # user success message
        sourceFlag = True # data can be processed by the feeder processing functions
    except:
        loadMessage33K.configure(text="X", fg="red")
        userMessage.configure(text="حدث خطأ اثناء معالجة ملف المصادر", fg="red") # user failure message
        sourceFlag = False # data will not be processed by the feeder processing functions

""""
Functionality: 
    Read data from loads excel file row by row, looks for loads data of each existing feeder and sources (in Feeder and Source classes)
Return:
    nothing!
"""
def import_loads():
    """ include the global variables related to transformers, if not included, they can't be modified inside the function """
    global loadFlag
    global loadMessageLoads # this message to indicate if the file is properly loaded
    global userMessage
    """ if current date exceeded the expiry date, the program will show error message and stops working """
    if not validate_date():
        userMessage.configure(text="هنالك خطأ في البرنامج, اتصل بالمصصم على الرقم 07701791983 ", fg="red")
        return
    """ load file depends on feeders (11 KV) and sources (33 KV), so the feeders and sources files must be uploaded first"""
    if not (feederFlag and sourceFlag):
        userMessage.configure(text="قم بتحميل جداول مغذيات (11 كف) و مصادر (33 كف) اولا ", fg="red")
        return
    """ Create constant variables instead of using the dictionary, make it cleaner and also easier to maintain in the future. """  
    LOAD = LOAD_NAMES["LOAD"]
    VOLTS = LOAD_NAMES["VOLTS"]
    NAME = LOAD_NAMES["FEEDER"]
    try:
        filename = filedialog.askopenfilename(initialdir = "/",title = "اختر ملف الاحمال",filetypes = (("Excel files","*.xls"),("all files","*.*")))
        loadFrame = pandas.read_excel(filename, sheet_name=0) # Create panda fram  reading excel file
    except:
        userMessage.configure(text="لم يتم تحميل ملف الاحمال", fg="red")
        loadFlag = False
        loadMessageLoads.configure(text="X", fg="red")
        return
    headers = loadFrame.columns.tolist() # Create a list contains all column header of the excel sheet
    """ Validate the headers of the excel sheet """
    if not validate_columns(LOAD_NAMES, headers):
        userMessage.configure(text="هنالك عدم مطابقة في عناوين ملف الاحمال", fg="red")
        transFlag = False
        return
    """ 
    Read the excel sheet (stored in pandas frame) row by row, and store the loads in Source and Feeder class objects
    """
    try:    
        for index, row in loadFrame.iterrows():
            name = str(row[NAME]).strip() # remove leading spaces from the feeder name
            if row[VOLTS] == "11 KV":
                feeder = Feeder.objectsDic.get(name, None) # check if the feeder already exist in the feeders list
                """ if feeder exist, add transformers data to it, if not, ignore it. """
                if feeder is not None:
                    feeder.load = row[LOAD]
            elif row[VOLTS] == "33 KV":
                source = Source.objectsDic.get(name, None) # check if the feeder already exist in the feeders list
                """ if feeder exist, add transformers data to it, if not, ignore it. """
                if source is not None:
                    source.load = row[LOAD]
            else:
                print(f"Feeder {row[NAME]} has wrong voltage field")
        loadMessageLoads.configure(text=f"{CHECK_MARK}", fg="green")
        userMessage.configure(text=f"تمت معالجة ملف الاحمال ", fg="green") # user success message
        loadFlag = True # data can be processed by the feeder processing functions
    except:
        loadMessageLoads.configure(text="X", fg="red")
        userMessage.configure(text="حدث خطأ اثناء تحميل ملف الاحمال", fg="red") # user failure message
        loadFlag = False # data will not be processed by the feeder processing functions

"""
Functionality:
    if 4 excel files of feeders, transformers, sources and loads are uploaded, and processed properly, 
    this method will create an excel file, and puts in it all the data required for the sources report.
Returns:
    Nothing !
"""
def export_sources_report():
    global userMessage
    """ if current date exceeded the expiry date, the program will show error message and stops working """
    if not validate_date():
        userMessage.configure(text="هنالك خطأ في البرنامج, اتصل بالمصصم على الرقم 07701791983 ", fg="red")
        return
    """
    First check whether the two excel files were uploaded and processed properly,
    if not, the method will stop and ask user to upload and process the proper files 
    """
    if not (feederFlag and sourceFlag and loadFlag):
        userMessage.configure(text="تأكد من تحميل الملفات بصورة صحيحة قبل محاولة تصدير التقرير", fg="red")
        return
    try:
        """ get a file name from user browsing box """
        filename = filedialog.asksaveasfilename(filetypes=(("Excel files", "*.xlsx"),("All files", "*.*") ))
        """if the user didn't specify a path, an error message will be displayed"""
        if filename is None or filename == "":
            userMessage.configure(text="لم يتم تحديد مسار ملف تقرير المصادر", fg="red")
            return
        """ create excel file workbook, and a worksheet, and customize the worksheet """
        workbook = xlsxwriter.Workbook(filename + ".xlsx", {'nan_inf_to_errors': True}) #{'nan_inf_to_errors': True} is the option to allow wirte float('nan') into excel cells
        worksheet = workbook.add_worksheet()
        worksheet.right_to_left() # make it arabic oriented
        worksheet.set_zoom(70) # the zoom will be 70%
        """ 
        Create cell style per each type of cells
        style (format) will be added per cell, because I found it eaiser to perform in XlsxWriter
        """
        titleCellFormat = workbook.add_format({'bold': True, 'font_size':14, 'align': 'center', 'valign':'vcenter', 'border':True, 'pattern':1, 'bg_color':'#d3d3d3'})
        seperatorCellFormat = workbook.add_format({'bold': True, 'font_size':14, 'align': 'center', 'valign':'vcenter', 'border':True, 'pattern':1, 'bg_color':'red'})
        logoCellFormat = workbook.add_format({'bold': True, 'font_size':18, 'align': 'center', 'valign':'vcenter', 'border':True})
        genericCellFormat = workbook.add_format({'align': 'center', 'valign':'vcenter', 'border':True})
        sumCellFormat = workbook.add_format({'bold': True, 'font_size':14, 'align': 'center', 'valign':'vcenter', 'border':True})
        """ set the width of columns """
        worksheet.set_column("A:A",20)
        worksheet.set_column("B:B",20)
        worksheet.set_column("C:C",13)        
        worksheet.set_column("D:D",18)
        worksheet.set_column("E:E",13)
        worksheet.set_column("F:F",20)
        worksheet.set_column("G:G",13)
        worksheet.set_column("H:J",12)
        worksheet.set_column("AC:AC",20)
        """
        Build title and log, which will be first 4 rows
        """
        """ 1st row for logo image, but I didn't load the image due to problem with pyinstaller --onefile """
        worksheet.merge_range("A1:AC1", "", genericCellFormat) 
        worksheet.set_row(0,210)
        worksheet.insert_image('A1', 'images\ministry.png', {'x_scale': 1.89, 'y_scale': 1.451})
        """ 2nd row for department title """
        worksheet.merge_range("A2:AC2", "مديرية توزيع كهرباء مركز نينوى", logoCellFormat)
        worksheet.set_row(1,40)
        """ 3rd and 4th rows for columns titles """
        worksheet.set_row(2,25)
        worksheet.set_row(3,25)
        for cellRange, text in (["A3:A4","محطات 132"],["B3:B4","مصادر 33 كف"],["C3:C4","حمل المصدر"],["D3:D4","محطات 33"],["E3:E4","جانب المدينة"],["F3:F4","مغذيات 11 كف"],["G3:G4","حمل المغذي"]):
            worksheet.merge_range(cellRange, text, titleCellFormat)
        worksheet.merge_range("H3:J3", "اطوال المغذيات - بالمتر", titleCellFormat)
        for cellRange, text in (["H4","ارضي"],["I4","هوائي"],["J4","الكلي"]):
            worksheet.write(cellRange, text, titleCellFormat)
        titleColumnIndex = 10 # transformers' columns titles start at the 11th column (first 11 columns taken for station, city side, feeder length, etc.)
        for cellRange, text in (["K3:P3","محولات صندوقية"],["Q3:V3","غرف محولات"],["W3:AB3","محولات هوائية"]):
            worksheet.merge_range(cellRange, text, titleCellFormat)
            for size in [100, 250, 400, 630, 1000, "اخرى"]:
                worksheet.write(3, titleColumnIndex, size, titleCellFormat)
                titleColumnIndex += 1
        worksheet.merge_range("AC3:AC4","مجموع المحولات", titleCellFormat)
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
        totalFeederLoads = 0
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
            columnIndex = 5 # first two columns are taken for station name, and city side
            for feeder in feedersList:
                for value in (feeder.name, feeder.load, feeder.cableLength, feeder.overLength, feeder.totalLength()):
                    worksheet.write(endRowIndex, columnIndex, value, genericCellFormat)
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
                totalFeederLoads += feeder.load
                totalCableLength += feeder.cableLength
                totalOverLength += feeder.overLength
                columnIndex = 5 # reset column index for each new feeder
                endRowIndex += 1 # end row index refer to next empty row
            worksheet.merge_range(startRowIndex,3,endRowIndex-1,3, name, genericCellFormat) # add the station in the first column, with height equal all feeder rows
            worksheet.merge_range(startRowIndex,4,endRowIndex-1,4, station.citySide, genericCellFormat) # add the city side in the first column, with height equal all feeder rows           
            worksheet.merge_range(endRowIndex,0,endRowIndex,28, "", seperatorCellFormat) # create an empty row, works as separation between stations
            endRowIndex += 1 # increase the row pointer to point to the next row after the empty one added.
            startRowIndex = endRowIndex # At the end of each new loop, the row start and end indexes should be equal
        """ finally, add the sumation row at the bottom of the sheet """
        columnIndex = 3
        totalCombinedLength = totalCableLength + totalOverLength
        for text in ["المجموع الكلي", "", totalFeeders, totalFeederLoads, totalCableLength, totalOverLength, totalCombinedLength]:
            worksheet.write(endRowIndex, columnIndex, text, sumCellFormat)
            columnIndex += 1
        for shape in ['kiosk', 'indoor', 'outdoor']:
            for size in ['100', '250', '400', '630', '1000', 'other']:
                totalTransCol = totalTransTypes[shape][size]
                worksheet.write(endRowIndex, columnIndex, totalTransCol, sumCellFormat)
                grandTransSum += totalTransCol
                columnIndex += 1
        worksheet.write(endRowIndex, 28, grandTransSum, sumCellFormat) 
        worksheet.set_row(endRowIndex, 40) # set the height of row, I couldn't do at the beginning with other formats becuase it uses a variable the its value couldn't be known at the beginning
        workbook.close() # finally save the excel file
        userMessage.configure(text=f"تم تصدير تقريرالمصادر ", fg="green") # user success message
    except:
        userMessage.configure(text="حدث خطأ اثناء تحميل تقرير المصادر", fg="red") # user message if any thing went wrong during executing the function
    return

"""
Functionality:
    if both excel files of feeders, and transformers are uploaded, and processed properly, 
    this method will create an excel file, and puts in it all the data required for the ministery report.
Returns:
    Nothing !
"""
def export_ministery_report():
    global userMessage
    """ if current date exceeded the expiry date, the program will show error message and stops working """
    if not validate_date():
        userMessage.configure(text="هنالك خطأ في البرنامج, اتصل بالمصصم على الرقم 07701791983 ", fg="red")
        return
    """
    First check whether the two excel files were uploaded and processed properly,
    if not, the method will stop and ask user to upload and process the proper files 
    """
    if not feederFlag or not transFlag:
        userMessage.configure(text="تأكد من تحميل الملفات بصورة صحيحة قبل محاولة تصدير التقرير", fg="red")
        return
    try:
        """ get a file name from user browsing box """
        filename = filedialog.asksaveasfilename(filetypes=(("Excel files", "*.xlsx"),("All files", "*.*") ))
        """if the user didn't specify a path, an error message will be displayed"""
        if filename is None or filename=="":
            userMessage.configure(text="لم يتم تحديد مسار ملف الوزارة", fg="red")
            return
        """ create excel file workbook, and a worksheet, and customize the worksheet """
        workbook = xlsxwriter.Workbook(filename + ".xlsx", {'nan_inf_to_errors': True}) #{'nan_inf_to_errors': True} is the option to allow wirte float('nan') into excel cells
        worksheet = workbook.add_worksheet()
        worksheet = workbook.add_worksheet()
        worksheet.right_to_left() # make it arabic oriented
        worksheet.set_zoom(70) # the zoom will be 70%
        """ 
        Create cell style per each type of cells
        style (format) will be added per cell, because I found it eaiser to perform in XlsxWriter
        """
        titleCellFormat = workbook.add_format({'bold': True, 'font_size':14, 'align': 'center', 'valign':'vcenter', 'border':True, 'pattern':1, 'bg_color':'#d3d3d3'})
        seperatorCellFormat = workbook.add_format({'bold': True, 'font_size':14, 'align': 'center', 'valign':'vcenter', 'border':True, 'pattern':1, 'bg_color':'red'})
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
        """ 1st row for logo image, but I didn't load the image due to problem with pyinstaller --onefile """
        worksheet.merge_range("A1:Y1", "", genericCellFormat) 
        worksheet.set_row(0,210)
        worksheet.insert_image('A1', 'images\ministry.png', {'x_scale': 1.451, 'y_scale': 1.451})
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
            worksheet.merge_range(endRowIndex,0,endRowIndex,24, "", seperatorCellFormat) # create an empty row, works as separation between stations
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
        userMessage.configure(text=f"تم تصدير تقرير الوزارة ", fg="green") # user success message
    except:
        userMessage.configure(text="حدث خطأ اثناء تصدير تقرير الوزارة", fg="red") # user message if any thing went wrong during executing the function
    return

"""
Functionality:
    if both excel files of feeders, and transformers are uploaded, and processed properly, 
    this method will create an excel file, and creates the transformers report
Returns:
    Nothing !
"""
def export_transformers_report():
    global userMessage
    """ if current date exceeded the expiry date, the program will show error message and stops working """
    if not validate_date():
        userMessage.configure(text="هنالك خطأ في البرنامج, اتصل بالمصصم على الرقم 07701791983 ", fg="red")
        return
    """
    First check whether the two excel files were uploaded and processed properly,
    if not, the method will stop and ask user to upload and process the proper files 
    """
    if not feederFlag or not transFlag:
        userMessage.configure(text="تأكد من تحميل الملفات بصورة صحيحة قبل محاولة تصدير التقرير", fg="red")
        return
    try:
        """ get a file name from user browsing box """
        filename = filedialog.asksaveasfilename(filetypes=(("Excel files", "*.xlsx"),("All files", "*.*") ))
        """if the user didn't specify a path, an error message will be displayed"""
        if filename is None or filename=="":
            userMessage.configure(text="لم يتم تحديد مسار تصدير ملف المحولات", fg="red") 
            return
        """ create excel file workbook, and a worksheet, and customize the worksheet """
        workbook = xlsxwriter.Workbook(filename + ".xlsx", {'nan_inf_to_errors': True}) #{'nan_inf_to_errors': True} is the option to allow wirte float('nan') into excel cells
        worksheet = workbook.add_worksheet()
        worksheet = workbook.add_worksheet()
        worksheet.right_to_left() # make it arabic oriented
        titleFormat = workbook.add_format({'align': 'center', 'valign':'vcenter', 'border':True, 'pattern':1, 'bg_color':'#d3d3d3'})
        cellFormat = workbook.add_format({'align': 'center', 'valign':'vcenter', 'border':True})
        worksheet.set_column("A:A",20)
        worksheet.set_column("C:C",45) 
        """ fill title row """
        worksheet.write("A1", "اسم المغذي", titleFormat)
        worksheet.write("B1", "النوع", titleFormat)
        worksheet.write("C1", "المحولات", titleFormat)
        """ build sheet row by row """
        rowIndex = 1
        for name, feeder in Feeder.objectsDic.items():
            transText = {"Cable":"", "Over":""}
            totalTrans = 0
            for size in ('100', '250', '400', '630', '1000','other'):
                idTransNum = feeder.trans['indoor'][size] + feeder.trans['kiosk'][size]
                odTransNum = feeder.trans['outdoor'][size]
                transText["Cable"] += f"{size}x{idTransNum} + "
                transText["Over"] += f"{size}x{odTransNum} + "
                totalTrans += odTransNum
            """ if no transformers in the overhead, then make the text empty, so it will not be added to excel """
            if totalTrans == 0:
                transText["Over"] = ""
            for cableType, transText in transText.items():
                if len(transText) > 0:
                    worksheet.write(rowIndex, 0, name, cellFormat)
                    worksheet.write(rowIndex, 1, cableType, cellFormat)
                    worksheet.write(rowIndex, 2, transText[:-2], cellFormat) # Remove last two char (" +") from the transformer text
                    rowIndex += 1
        workbook.close() # finally save the excel file
        userMessage.configure(text=f"تم تصدير تقرير عدد المحولات ", fg="green") # user success message
    except:
        userMessage.configure(text="حدث خطأ اثناء تصدير تقرير المحولات", fg="red") # user message if any thing went wrong during executing the function

"""
Functionality:
    If the time exceeded predifined date (expiry date), the program will not work
return: boolean
    True: If the program did not exceeded the expiry date
    False: If the program is not valid anymore (exceeded expiry date)
"""
def validate_date():
    currentDate = datetime.now()
    expiryDate = datetime.strptime("1/6/2020 4:00", "%d/%m/%Y %H:%M")
    return expiryDate > currentDate

"""
Functionality:
    main function that provides an interface for running the program
return: 
    Nothing
"""
def main():
    """ messages accessed in other parts of the program"""
    global userMessage
    global loadMessage11K
    global loadMessage33K
    global loadMessageLoads
    global loadMessageTrans
    """ constructing the GUI """
    window = tkinter.Tk()
    window.title("GIS Reports V1.0")
    window.geometry("1000x800")
    """ GIS logo """
    logoFrame = Frame(window)
    gisLable1 =  Label(logoFrame, text="قسم التخطيط", fg="navy", font=("Helvetica", 20))
    gisLable1.pack(side=RIGHT, padx=10, pady=10)
    logoImage = PhotoImage(file = r"images\logo.png").subsample(10, 10) # create photo and resize it (with subsample)
    gisLogo = Label(logoFrame, image = logoImage)
    gisLogo.pack(side=RIGHT, padx=5, pady=5)
    gisLable1 =  Label(logoFrame, text="شعبة المعلوماتية", fg="navy", font=("Helvetica", 20))
    gisLable1.pack(side=RIGHT, padx=10, pady=10)
    logoFrame.pack()
    """ importing files """
    importGroup = LabelFrame(window, text="    تحميل الملفات    ", padx=20, pady=10, labelanchor=NE)
    importGroup.pack(pady=15)
    openImage = PhotoImage(file = r"images\open.png").subsample(5, 5) # create photo and resize it (with subsample)
    """ left subframe """
    leftSubFrame = Frame(importGroup)
    leftSubFrame.pack(side=LEFT, padx=5, pady=5)
    group33K = Frame(leftSubFrame)
    group33K.pack()
    loadMessage33K = Label(group33K, text="X", fg="red", font=("Helvetica", 16))
    loadMessage33K.pack(side=LEFT) 
    feeder33Button = Button(group33K, text="   مصادر 33 كف   ", image = openImage, compound = 'right', command=import_sources, cursor="hand2", font=("Helvetica", 14))
    feeder33Button.pack(side=RIGHT, padx=10, pady=5, ipadx=15, ipady=5)
    groupLoad = Frame(leftSubFrame)
    groupLoad.pack()
    loadMessageLoads = Label(groupLoad, text="X", fg="red", font=("Helvetica", 16))
    loadMessageLoads.pack(side=LEFT) 
    loadsButton = Button(groupLoad, text="    جدول احمال      ", image = openImage, compound = 'right', command=import_loads, cursor="hand2", font=("Helvetica", 14))
    loadsButton.pack(side=RIGHT, padx=10, pady=5, ipadx=15, ipady=5)
    """ right subframe """
    rightSubFrame = Frame(importGroup)
    rightSubFrame.pack(side=RIGHT, padx=5, pady=5)
    group11K = Frame(rightSubFrame)
    group11K.pack()
    loadMessage11K = Label(group11K, text="X", fg="red", font=("Helvetica", 16))
    loadMessage11K.pack(side=LEFT)
    feeder11Button = Button(group11K, text="   مغذيات 11 كف   ", image = openImage, compound = 'right', command=import_feeders, cursor="hand2", font=("Helvetica", 14))
    feeder11Button.pack(side=RIGHT, padx=10, pady=5, ipadx=15, ipady=5)
    groupTrans = Frame(rightSubFrame)
    groupTrans.pack()
    loadMessageTrans = Label(groupTrans, text="X", fg="red", font=("Helvetica", 16))
    loadMessageTrans.pack(side=LEFT)
    transButton = Button(groupTrans, text="  جدول المحولات    ", image = openImage, compound = 'right', command=import_transformers, cursor="hand2", font=("Helvetica", 14))
    transButton.pack(sid=RIGHT, padx=10, pady=5, ipadx=15, ipady=5)
    """ save files """
    saveGroup = LabelFrame(window, text="    تصدير النتائج    ", padx=20, pady=10, labelanchor=NE)
    saveGroup.pack(pady=15)
    saveImage = PhotoImage(file = r"images\save.png").subsample(5, 5) # create photo and resize it
    exportMinistry = Button(saveGroup, text="  تقرير الوزارة", image = saveImage, compound = 'left', command=export_ministery_report, cursor="hand2", font=("Helvetica", 14))
    exportTrans = Button(saveGroup, text="  تقرير المحولات", image = saveImage, compound = 'left', command=export_transformers_report, cursor="hand2", font=("Helvetica", 14))
    export33Kv = Button(saveGroup, text="   تتقرير المصادر  ", image = saveImage, compound = 'left', command=export_sources_report, cursor="hand2", font=("Helvetica", 14))
    exportTrans.pack(side=RIGHT, padx=10, pady=10, ipadx=15, ipady=7)
    exportMinistry.pack(side=RIGHT, padx=10, pady=10, ipadx=15, ipady=7)
    export33Kv.pack(side=RIGHT, padx=10, pady=10, ipadx=15, ipady=7)
    """ User message """
    messageGroup = LabelFrame(window, text="    رسائل المستخدم    ", padx=20, pady=5, labelanchor=NE)
    messageGroup.pack(padx=7, pady=7)
    userMessage = Label(messageGroup, text=" مرحبا بك في برنامج تقارير قسم المعلوماتية ", fg="green", font=("Helvetica", 14))
    userMessage.pack(padx=7, pady=7)
    """ Exit Button """
    exitImage = PhotoImage(file = r"images\exit.png").subsample(6, 6) # create photo and resize it
    button = Button(window, text = "  خروج  ", image = exitImage, compound = 'left', command = window.destroy, cursor="hand2", fg="navy", font=("Helvetica", 14)) # close the program window
    button.pack(side=LEFT, padx=20, pady=20, ipadx=10, ipady=10 )

    window.mainloop()

if __name__ == '__main__':
    main()
