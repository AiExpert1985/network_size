"""
this is the styling version
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

class Station11K:
    """
    Each station11k will contains its info and multiple unique feeder objects
    """
    stationsDic = {}
    def __init__(self, name, citySide):
        self.name = name
        self.citySide = citySide
        self.feedersList = []
        Station11K.stationsDic[name] = self
    def addFeeder(self, feeder):
        self.feedersList.append(feeder)

class Feeder:
    """
    Each feeder object contains all info related to the feeder
    """
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

def validate_columns(desiredHeaders, fileHeaders):
    """
    Functionality: Compare the dsired column headers (stored in the program) with the column headers read from excel sheet
    Parameters:
        (1) dictionary: the desired columns headers (stored in the program)
        (2) list: headers read from the excel file
    return: boolean
        True: If stored headers, and headers read from the excel file are the same
        False: If stored headers, and headers read from the excel file are not the same
    """
    for key, value in desiredHeaders.items():
        if value not in fileHeaders:
            return False
    return True

# extract all the info inside feeder excel sheet, and store it in Feeder and Station11k classes          
def import_feeders():
    """"
    Functionality: 
        Read data from feeder excel file, put the data in both Stations11k, and Feeder classes.
        only the data of feeders which there status is 'بالعمل' are read, others will be neglected
    Return:
        nothing!
    """
    # include the global variables related to feeder, if not included, they can't be modified inside the function
    global feedFrame
    global feederFlag
    # Create constant variables instead of using the dictionary, make it cleaner and also easier to maintain in the future
    FEEDER_FEEDER = FEEDER_NAMES["FEEDER"]
    FEEDER_STATION = FEEDER_NAMES["STATION"]
    FEEDER_CITYSIDE = FEEDER_NAMES["CITYSIDE"]
    FEEDER_TYPE = FEEDER_NAMES["TYPE"]
    FEEDER_LENGTH = FEEDER_NAMES["LENGTH"]
    FEEDER_STATUS = FEEDER_NAMES["STATUS"]
    FEEDER_NUMBER = FEEDER_NAMES["NUMBER"]
    # Upload excel file contain the feeders data
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
    """ Read the excel sheet (stored in pandas frame) row by row, and store result in Station11k, and Feeder classes
        rows will be neglected if the status is not (بالعمل) """
    try:
        for index, row in feedFrame.iterrows():
            if row[FEEDER_STATUS] == "بالعمل":
                feederName = str(row[FEEDER_FEEDER]).strip() # remove leading spaces from the feeder name
                """ check if the feeder was previously read, 
                    if yes, then the data will be read and stored in the same feeder, 
                    if not, a new feeder will be created, and then data stored in it """
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
        feederFlag = True
    except:
        userMessage.configure(text="حدث خطأ اثناء معالجة بيانات ملف المغذيات", fg="red") # Display failure message to user
        feederFlag = False # data will not be processed by the feeder processing functions
 
# import transformers info 
def import_transformers():
    global transFrame
    global transFlag
    TRANS_FEEDER = TRANS_NAMES["FEEDER"]
    TRANS_SIZE = TRANS_NAMES["SIZE"]
    TRANS_TYPE = TRANS_NAMES["TYPE"]
    TRANS_STATUS = TRANS_NAMES["STATUS"]
    try:
        filename = filedialog.askopenfilename(initialdir = "/",title = "اختر ملف المحولات",filetypes = (("Excel files","*.xls"),("all files","*.*")))
        transFrame = pandas.read_excel(filename,sheet_name=0)
    except:
        userMessage.configure(text="خطأ اثناء تحميل ملف المحولات", fg="red")
        transFlag = False
        return
    headers = transFrame.columns.tolist()
    if not validate_columns(TRANS_NAMES,headers):
        userMessage.configure(text="هنالك عدم مطابقة في عناوين الاعمدة في ملف المحولات", fg="red")
        transFlag = False
        return
    try:        
        for index, row in transFrame.iterrows():
            if row[TRANS_STATUS] == "good":
                name = str(row[TRANS_FEEDER]).strip()
                feeder = Feeder.objectsDic.get(name, None)
                if feeder is not None:
                    transType = row[TRANS_TYPE]
                    transSize = row[TRANS_SIZE]
                    if transSize in ['100','250','400','630','1000'] and transType in ['indoor', 'outdoor', 'kiosk']:
                        feeder.trans[transType][transSize] += 1
                    else:
                        if transType in ['indoor', 'outdoor', 'kiosk']:
                            feeder.trans[transType]['other'] += 1
        userMessage.configure(text=f"تمت معالجة ملف المحولات {CHECK_MARK}", fg="green")
        transFlag = True
    except:
        userMessage.configure(text="حدث خطأ اثناء معالجة بيانات ملف المحولات", fg="red")
        transFlag = False
    
def export_ministery_report():
    if feederFlag and transFlag:
        exportExcel()
        userMessage.configure(text=f"تم تصدير تقرير الوزارة {CHECK_MARK}", fg="green")
    else:
        userMessage.configure(text="تأكد من تحميل الملفات بصورة صحيحة قبل محاولة تصديرها", fg="red")
        
def export_transformers_text():
    if feederFlag and transFlag:
        exportExcel()
        userMessage.configure(text=f"تم تصدير تقرير عدد المحولات {CHECK_MARK}", fg="green")
    else:
        userMessage.configure(text="تأكد من تحميل الملفات بصورة صحيحة قبل محاولة تصديرها", fg="red")

def exportExcel():
    try:
        # get the file name
        filename = filedialog.asksaveasfilename(filetypes=(("Excel files", "*.xlsx"),("All files", "*.*") ))
        if filename is None: # asksaveasfile return `None` if dialog closed with "cancel".
            return
        # create workboo, and work sheet, and customize worksheet direction and size
        workbook = xlsxwriter.Workbook(filename + ".xlsx")
        worksheet = workbook.add_worksheet()
        worksheet.right_to_left()
        worksheet.set_zoom(70)
        start_row = 4
        end_row = 4
        # formats of cell // here we format per cell, not per column
        feeder_cell_format = workbook.add_format({'valign':'vcenter', 'border':True})
        title_cell_format = workbook.add_format({'align': 'center', 'valign':'vcenter', 'border':True, 'pattern':1, 'bg_color':'#d3d3d3'})
        cell_format = workbook.add_format({'align': 'center', 'valign':'vcenter', 'border':True})
        sumFormat = workbook.add_format({'bold': True, 'font_size':14, 'align': 'center', 'valign':'vcenter', 'border':True})
        # titles of the columns
        worksheet.insert_image('A1', 'logo.png', {'x_scale': 1.451, 'y_scale': 1.451})
        worksheet.merge_range("A1:Y1", "GIS logo", cell_format)
        worksheet.set_row(0,210)
        worksheet.merge_range("A2:Y2", "مديرية توزيع كهرباء مركز نينوى", cell_format)
        worksheet.set_row(1,25)
        worksheet.merge_range("A3:A4", "اسم المحطة", title_cell_format)
        worksheet.merge_range("B3:B4", "جانب المدينة", title_cell_format)
        worksheet.merge_range("C3:C4","اسم المغذي", title_cell_format)
        worksheet.merge_range("D3:F3","اطوال المغذيات (متر)", title_cell_format)
        worksheet.write("D4", "ارضي", title_cell_format)
        worksheet.write("E4", "هوائي", title_cell_format)
        worksheet.write("F4", "الكلي", title_cell_format)
        worksheet.merge_range("G3:L3","محولات صندوقية", title_cell_format)
        worksheet.merge_range("M3:R3","غرف محولات", title_cell_format)
        worksheet.merge_range("S3:X3","محولات هوائية", title_cell_format)
        # add titles to transformers sizes
        titleColumnIndex = 6
        for i in range(3):
            for title in [100, 250, 400, 630, 1000, "اخرى"]:
                worksheet.write(3, titleColumnIndex, title, title_cell_format)
                titleColumnIndex += 1
        worksheet.merge_range("Y3:Y4","مجموع المحولات", title_cell_format)
        # initiate counters, will be the last row in the sheet
        totalFeeders = 0
        totalCableLength = 0
        totalOverLength = 0
        totalCombinedLength = 0
        totalTrans = {
                "kiosk": {"100":0, "250":0, "400":0, "630":0, "1000":0, "other":0},
                "indoor": {"100":0, "250":0, "400":0, "630":0, "1000":0, "other":0},
                "outdoor": {"100":0, "250":0, "400":0, "630":0, "1000":0, "other":0}
                }
        totalCombinedTrans = 0
        # intersts data as rows
        for name, station in Station11K.stationsDic.items():
            feedersList = station.feedersList
            # sorting feeders inside a station according to their numbers
            feedersList.sort(key=lambda x: x.number, reverse=False)
            for feeder in feedersList:
                worksheet.write(end_row, 2, feeder.name, feeder_cell_format)
                worksheet.write(end_row, 3, feeder.cableLength, cell_format)
                worksheet.write(end_row, 4, feeder.overLength, cell_format)
                worksheet.write(end_row, 5, feeder.totalLength(), cell_format)
                transColumnIndex = 6
                transTotal = 0
                colors = {'kiosk':'#fef200', 'indoor':'#75d86a', 'outdoor':'#4dc3ea'}
                for shape in ['kiosk', 'indoor', 'outdoor']:
                    color = colors[shape]
                    transCellFormat = workbook.add_format({'align': 'center', 'valign':'vcenter', 'border':True, 'pattern':1, 'bg_color':color})
                    for size in ['100', '250', '400', '630', '1000', 'other']:
                        transSum = feeder.trans[shape][size]
                        transTotal += transSum
                        worksheet.write(end_row, transColumnIndex, transSum, transCellFormat)
                        transColumnIndex += 1
                        totalTrans[shape][size] += transSum
                worksheet.write(end_row, 24, transTotal, cell_format)
                end_row += 1
                totalFeeders += 1
                totalCableLength += feeder.cableLength
                totalOverLength += feeder.overLength
            worksheet.merge_range(start_row,0,end_row-1,0, name, cell_format)
            worksheet.merge_range(start_row,1,end_row-1,1, station.citySide, cell_format)
            worksheet.merge_range(end_row,0,end_row,24, "", title_cell_format)
            end_row += 1
            start_row = end_row
        totalCombinedLength = totalCableLength + totalOverLength
        worksheet.write(end_row, 0, "المجموع الكلي", sumFormat)
        worksheet.write(end_row, 2, totalFeeders, sumFormat)
        worksheet.write(end_row, 3, totalCableLength, sumFormat)
        worksheet.write(end_row, 4, totalOverLength, sumFormat)
        worksheet.write(end_row, 5, totalCombinedLength, sumFormat)
        grandColIndex = 6
        for shape in ['kiosk', 'indoor', 'outdoor']:
            for size in ['100', '250', '400', '630', '1000', 'other']:
                transColSum = totalTrans[shape][size]
                worksheet.write(end_row, grandColIndex, transColSum, sumFormat)
                grandColIndex += 1
                totalCombinedTrans += transColSum
        worksheet.write(end_row, 24, totalCombinedTrans, sumFormat)
        worksheet.set_column("A:A",18)
        worksheet.set_column("B:B",12)        
        worksheet.set_column("C:C",20)
        worksheet.set_column("D:F",12)
        worksheet.set_column("Y:Y",15)
        worksheet.set_row(end_row, 40)
        workbook.close()
    except:
        userMessage.configure(text="حدث خطأ اثناء تصدير الملف", fg="red")

def transTextDic():
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
