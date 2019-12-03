###### this is the styling version

import pandas
import tkinter
from tkinter import Button
from tkinter import filedialog
from tkinter import Label
import xlsxwriter

# CONSTANTS
# titles dictionaries: Columns names for feeder, and transformer sheets, program will look for these titles in the uploaded feeder and transformer excel sheets
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
# this constant is used for displaying check mark sign for user messages
CHECK_MARK = u'\u2713'

# global variables
# these variables are updated in the functions, but they must be declared as global at the begining of functions
feedFrame = None # contains the frame read from feeders excel
transFrame = None #contains the frame read from transformers excel
userMessage = None # contains message displayed 
f_flag = False # will be true if the the feeder file uploaded and processed successfully
t_flag = False # will be true if the the transformers file uploaded and processed successfully

class Station11K:
    stationsDic = {}
    def __init__(self, name):
        self.name = name
        self.feedersList = []
        Station11K.stationsDic[name] = self
    def addFeeder(self, feeder):
        self.feedersList.append(feeder)

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

# validate if the columns titles in the sheet are the same as titles dictionaries
def validate_columns(NAMES_LIST, columnsHeaders):
    for key, value in NAMES_LIST.items():
        if value not in columnsHeaders:
            return False
    return True

# import excel sheet contains feeders info            
def import_feeders():
    global feedFrame
    global f_flag
    FEEDER_FEEDER = FEEDER_NAMES["FEEDER"]
    FEEDER_STATION = FEEDER_NAMES["STATION"]
    FEEDER_CITYSIDE = FEEDER_NAMES["CITYSIDE"]
    FEEDER_TYPE = FEEDER_NAMES["TYPE"]
    FEEDER_LENGTH = FEEDER_NAMES["LENGTH"]
    FEEDER_STATUS = FEEDER_NAMES["STATUS"]
    FEEDER_NUMBER = FEEDER_NAMES["NUMBER"]
    try:
        filename = filedialog.askopenfilename(initialdir = "/",title = "اختر ملف المغذيات",filetypes = (("Excel files","*.xls"),("all files","*.*")))
        feedFrame = pandas.read_excel(filename,sheet_name=0)
    except:
        userMessage.configure(text="خطأ اثناء تحميل ملف المغذيات", fg="red")
        f_flag = False
        return
    headers = feedFrame.columns.tolist()
    if not validate_columns(FEEDER_NAMES,headers):
        userMessage.configure(text="هنالك عدم مطابقة في عناوين الاعمدة في ملف المغذيات", fg="red")
        f_flag = False
        return
    try:
        for index, row in feedFrame.iterrows():
            if row[FEEDER_STATUS] == "بالعمل":
                feederName = str(row[FEEDER_FEEDER]).strip()
                feeder = Feeder.objectsDic.get(feederName, None)
                if feeder is None:
                    feeder = Feeder(feederName)
                    feeder.number = row[FEEDER_NUMBER]
                    feeder.citySide = row[FEEDER_CITYSIDE]
                    stationName = row[FEEDER_STATION]
                    station = Station11K.stationsDic.get(stationName, None)
                    if station is None:
                        station = Station11K(stationName)
                    station.addFeeder(feeder)
                if row[FEEDER_TYPE] == "overhead":
                    feeder.overLength = round(row[FEEDER_LENGTH],2) # keep only two digits after the dot
                elif row[FEEDER_TYPE] == "cable":
                    feeder.cableLength = round(row[FEEDER_LENGTH],2) # keep only two digits after the dot
                else:
                    print(f"Feeder {row[FEEDER_FEEDER]} has ({row[FEEDER_TYPE]}) type field")
        userMessage.configure(text=f"تمت معالجة ملف المغذيات {CHECK_MARK}", fg="green")
        f_flag = True
    except:
        userMessage.configure(text="حدث خطأ اثناء معالجة بيانات ملف المغذيات", fg="red")
        f_flag = False
 
# import transformers info 
def import_transformers():
    global transFrame
    global t_flag
    TRANS_FEEDER = TRANS_NAMES["FEEDER"]
    TRANS_SIZE = TRANS_NAMES["SIZE"]
    TRANS_TYPE = TRANS_NAMES["TYPE"]
    TRANS_STATUS = TRANS_NAMES["STATUS"]
    try:
        filename = filedialog.askopenfilename(initialdir = "/",title = "اختر ملف المحولات",filetypes = (("Excel files","*.xls"),("all files","*.*")))
        transFrame = pandas.read_excel(filename,sheet_name=0)
    except:
        userMessage.configure(text="خطأ اثناء تحميل ملف المحولات", fg="red")
        t_flag = False
        return
    headers = transFrame.columns.tolist()
    if not validate_columns(TRANS_NAMES,headers):
        userMessage.configure(text="هنالك عدم مطابقة في عناوين الاعمدة في ملف المحولات", fg="red")
        t_flag = False
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
        t_flag = True
    except:
        userMessage.configure(text="حدث خطأ اثناء معالجة بيانات ملف المحولات", fg="red")
        t_flag = False
    
def export_ministery_report():
    if f_flag and t_flag:
        exportExcel()
        userMessage.configure(text=f"تم تصدير تقرير الوزارة {CHECK_MARK}", fg="green")
    else:
        userMessage.configure(text="تأكد من تحميل الملفات بصورة صحيحة قبل محاولة تصديرها", fg="red")
        
def export_transformers_text():
    if f_flag and t_flag:
        exportExcel()
        userMessage.configure(text=f"تم تصدير تقرير عدد المحولات {CHECK_MARK}", fg="green")
    else:
        userMessage.configure(text="تأكد من تحميل الملفات بصورة صحيحة قبل محاولة تصديرها", fg="red")

def exportExcel():
    try:
        filename = filedialog.asksaveasfilename(filetypes=(("Excel files", "*.xlsx"),("All files", "*.*") ))
        if filename is None: # asksaveasfile return `None` if dialog closed with "cancel".
            return
        workbook = xlsxwriter.Workbook(filename + ".xlsx")
        worksheet = workbook.add_worksheet()
        worksheet.right_to_left()
        start_row = 4
        end_row = 4
        # formats of cell // here we format per cell, not per column
        feeder_cell_format = workbook.add_format({'valign':'vcenter', 'border':True})
        title_cell_format = workbook.add_format({'align': 'center', 'valign':'vcenter', 'border':True, 'pattern':1, 'bg_color':'#d3d3d3'})
        cell_format = workbook.add_format({'align': 'center', 'valign':'vcenter', 'border':True})
        # titles of the columns
        worksheet.merge_range("A1:Y1", "GIS logo", cell_format)
        worksheet.set_row(0,100)
        worksheet.merge_range("A2:Y2", "مديرية توزيع كهرباء مركز نينوى", cell_format)
        worksheet.merge_range("A3:A4", "اسم المحطة", title_cell_format)
        worksheet.merge_range("B3:B4", "جانب المدينة", title_cell_format)
        worksheet.merge_range("C3:C4","اسم المغذي", title_cell_format)
        worksheet.merge_range("D3:F3","اطوال المغذيات (متر)", title_cell_format)
        worksheet.write("D4", "ارضي", title_cell_format)
        worksheet.write("E4", "هوائي", title_cell_format)
        worksheet.write("F4", "الكلي", title_cell_format)
        worksheet.merge_range("G3:L3","صندوقية", title_cell_format)
        worksheet.merge_range("M3:R3","غرف", title_cell_format)
        worksheet.merge_range("S3:X3","هوائية", title_cell_format)
        # add titles to transformers sizes
        col_index = 6
        for i in range(3):
            for title in [100, 250, 400, 630, 1000, "اخرى"]:
                worksheet.write(3, col_index, title, title_cell_format)
                col_index += 1
        worksheet.merge_range("Y3:Y4","مجموع المحولات", title_cell_format)
        # intersts data as rows
        for name, station in Station11K.stationsDic.items():
            feedersList = station.feedersList
            # sorting feeders inside a station according to their numbers
            feedersList.sort(key=lambda x: x.number, reverse=False)
            for feeder in feedersList:
                worksheet.write(end_row, 1, feeder.name, feeder_cell_format)
                worksheet.write(end_row, 2, feeder.cableLength, cell_format)
                end_row += 1
            worksheet.merge_range(start_row,0,end_row-1,0, name, cell_format)
            start_row = end_row
        worksheet.set_column(0,1,20)
        worksheet.set_column(2,5,15)
        worksheet.set_column(24,24,15)
        workbook.close()
    except:
        userMessage.configure(text="حدث خطأ اثناء تصدير الملف", fg="red")

def ministeryReportDic():
    feeders, stations, overLengths, cableLengths, totalLengths, citySides  = [],[],[],[],[],[]
    kio_100, kio_250, kio_400, kio_630, kio_1000, kio_other = [],[],[],[],[],[]
    id_100, id_250, id_400, id_630, id_1000, id_other = [],[],[],[],[],[]
    od_100, od_250, od_400, od_630, od_1000, od_other = [],[],[],[],[],[]
    for name, station in Station11K.stationsDic.items():
        feedersList = station.feedersList
        # sorting feeders inside a station according to their numbers
        feedersList.sort(key=lambda x: x.number, reverse=False)
        for feeder in feedersList:
            stations.append(station.name)
            feeders.append(feeder.name)
            overLengths.append(feeder.overLength)
            cableLengths.append(feeder.cableLength)
            totalLengths.append(feeder.totalLength())
            citySides.append(feeder.citySide)
            kio_100.append(feeder.trans['kiosk']['100'])
            kio_250.append(feeder.trans['kiosk']['250'])
            kio_400.append(feeder.trans['kiosk']['400'])
            kio_630.append(feeder.trans['kiosk']['630'])
            kio_1000.append(feeder.trans['kiosk']['1000'])
            kio_other.append(feeder.trans['kiosk']['other'])
            id_100.append(feeder.trans['indoor']['100'])
            id_250.append(feeder.trans['indoor']['250'])
            id_400.append(feeder.trans['indoor']['400'])
            id_630.append(feeder.trans['indoor']['630'])
            id_1000.append(feeder.trans['indoor']['1000'])
            id_other.append(feeder.trans['indoor']['other'])
            od_100.append(feeder.trans['outdoor']['100'])
            od_250.append(feeder.trans['outdoor']['250'])
            od_400.append(feeder.trans['outdoor']['400'])
            od_630.append(feeder.trans['outdoor']['630'])
            od_1000.append(feeder.trans['outdoor']['1000'])
            od_other.append(feeder.trans['outdoor']['other'])
    return {"stations":stations, "feeders":feeders, "overLengths": overLengths, "cableLengths":cableLengths, "totalLengths":totalLengths, "citySides":citySides,
            "kio_100":kio_100, "kio_250":kio_250, "kio_400":kio_400, "kio_630":kio_630, "kio_1000":kio_1000, "kio_other":kio_other, 
            "id_100":id_100, "id_250":id_250, "id_400":id_400, "id_630":id_630, "id_1000":id_1000, "id_other":id_other, 
            "od_100":od_100, "od_250":od_250, "od_400":od_400, "od_630":od_630, "od_1000":od_1000, "od_other":od_other}

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

if __name__ == '__main__':
    main()
