import pytz
import os
import sys
import re
import math
import unicodedata
from pathlib import Path
from datetime import datetime, date, timedelta
from ics import Calendar, Event
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.pagesizes import LETTER, A4, landscape
from reportlab.lib.units import inch
from reportlab.lib.colors import blue, green, black, red, pink, gray, brown, purple, orange, yellow
from reportlab.pdfbase import pdfmetrics  
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import registerFontFamily
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, Image, Spacer, Frame
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER

startdate = date(1990,1,1)
datecal = datetime.now()
calfont = "Georgia"
weekreps = []
columsmatrixreport = 3
rowsmatrixreport = 4
styles = getSampleStyleSheet()
#styles.list()
titleStyle = ParagraphStyle('hea', parent=styles['Normal'], fontSize = 12, textColor = black, alignment=TA_CENTER, leading = 8)
headerStyle = ParagraphStyle('hea', parent=styles['Normal'], fontSize = 12, textColor = orange, alignment=TA_CENTER, leading = 8)
weeksumStyle = ParagraphStyle('sum', parent=styles['Normal'], fontName = "ArialBold", fontSize = 12, textColor = green, leading = 8)
weeklocStyle = ParagraphStyle('loc', parent=styles['Normal'], fontName = "ArialItalic", fontSize = 9, textColor = blue, leading = 8)
weekdesStyle = ParagraphStyle('des', parent=styles['Normal'], fontName = "Arial", fontSize = 10, spaceAfter = 4, textColor = purple, leading = 8)
weektimStyle = ParagraphStyle('tim', parent=styles['Normal'], fontName = "Arial", fontSize = 9, spaceBefore = 4, spaceAfter = 0, textColor = red, leading = 8)
matrixsumheadingStyle = ParagraphStyle('sum', parent=styles['Normal'], fontName = calfont + "Bold", fontSize = 12, spaceBefore = 0, spaceAfter = 0, textColor = green, alignment=TA_CENTER, leading = 8)
matrixsumStyle = ParagraphStyle('sum', parent=styles['Normal'], fontName = calfont + "Bold", fontSize = 10, spaceBefore = 0, spaceAfter = 1, textColor = green, alignment=TA_CENTER, leading = 8)
matrixdesStyle = ParagraphStyle('des', parent=styles['Normal'], fontName = calfont, fontSize = 8, spaceBefore = 1, spaceAfter = 2, textColor = purple, alignment=TA_CENTER, leading = 8)
matrixtimlocStyle = ParagraphStyle('tim', parent=styles['Normal'], fontName = calfont, fontSize = 8, spaceBefore = 3, spaceAfter = 1, textColor = red, alignment=TA_CENTER, leading = 8)
title = Paragraph("Juli 2024", titleStyle)
weekdaynames = ["Maandag","Dinsdag","Woensdag","Donderdag","Vrijdag","Zaterdag","Zondag"]
monthnames = ["Januari","Februari","Maart","April","Mei","Juni","Juli","Augustus", "September","Oktober","November","December"]
weekStyle = [
('GRID',(1,1),(0,-1),3,green),
('BOX',(0,0),(1,-1),5,red),
('LINEABOVE',(1,2),(-2,2),1,blue),
('LINEBEFORE',(2,1),(2,-2),1,pink),
('FONTSIZE', (0, 1), (-1, 1), 10),
('VALIGN',(0,0),(3,0),'BOTTOM'),
('ALIGN',(0,0),(3,1),'CENTER')
]
matrixStyle = [
('BOX',(0,0),(-1,-1),3,red),
('VALIGN',(0,0),(-1,-1),'TOP'),
('ALIGN',(0,0),(3,1),'CENTER')
]
styles["Normal"].alignment = TA_LEFT
styles["Normal"].borderColor = red
styles["Normal"].textColor = green
styles["Normal"].fontSize = 8
    
class WeekReport:
    h0 =  [[] for _ in range(7)]
    p0 =  [[] for _ in range(7)]
    h1 =  [[] for _ in range(7)]
    p1 =  [[] for _ in range(7)]

    def append0_Paragraph(self, wkd, paragraph, style):
        textpar = Paragraph(paragraph, style)
        self.p0[wkd].append(textpar)
    
    def append0_Header(self, wkd, header, style):
        headerpar = Paragraph(header, style)
        self.h0[wkd].append(headerpar)
        
    def append1_Paragraph(self, wkd, paragraph, style):
        textpar = Paragraph(paragraph, style)
        self.p1[wkd].append(textpar)
    
    def append1_Header(self, wkd, header, style):
        headerpar = Paragraph(header, style)
        self.h1[wkd].append(headerpar)

    def clear(self):
        for i in range(7):
            while len(self.h0[i]) > 0:
                self.h0[i].pop()
            while len(self.p0[i]) > 0:
                self.p0[i].pop()
            while len(self.h1[i]) > 0:
                self.h1[i].pop()
            while len(self.p1[i]) > 0:
                self.p1[i].pop()

class MatrixReport:
    h = [[0 for i in range(columsmatrixreport)] for j in range(rowsmatrixreport)] 
    p = [[0 for i in range(columsmatrixreport)] for j in range(rowsmatrixreport)] 
 
    def clear(self):
        for r in self.h:
            for c in r:
                try:
                    c.pop()
                except IndexError:
                    print(c)
        for r in self.p:
            for c in r:
                try:
                    c.pop()
                except IndexError:
                    print(c)
    
class FamilienetEvent:
    def __init__(self, description, summary, weekday, weeknr, day, location, starttime, endtime, dayyear, month, alarm):
        self.description = description
        self.summary = summary
        self.weekday = weekday
        self.weeknr = weeknr
        self.day = day
        self.location = location
        self.starttime = starttime
        self.endtime = endtime
        self.dayyear = dayyear
        self.month = month
        self.alarm = alarm
        
def leapMonth(year, month):
    if month == 1 or month == 3 or month == 5 or month == 7 or month == 8 or month == 10 or month == 12:
        return 31
    if month == 2:
        if (year % 400) == 0:
            return 29
        elif (year % 100) == 0:
            return 28
        elif (year % 400) == 0:
             return 29
        else:
            return 28
    if month == 4 or month == 6 or month == 9 or month == 11:
        return 30
       
def weekDay(year, month, day):
    offset = [0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334]
    afterFeb = 1
    if month > 2: afterFeb = 0
    aux = year - 1700 - afterFeb
    dayOfWeek  = 5
    dayOfWeek += (aux + afterFeb) * 365                  
    dayOfWeek += aux / 4 - aux / 100 + (aux + 100) / 400     
    dayOfWeek += offset[month - 1] + (day - 1)               
    dayOfWeek %= 7
    return round(dayOfWeek)
    
def lookupimage(imgcode):
    newcalendar = [[] for _ in range(200)]
    newcalendar[171] = "MiniPrideWalk12juni-2.jpg"
    newcalendar[191] = "6617f09c1049a.jpg"
    newcalendar[190] = "keukenhof-lisse-molen.jpg"
    img = newcalendar[int(imgcode)]
    I = Image(img)
    I.drawHeight = 0.95*inch
    I.drawWidth = 0.95*inch
    I.hAlign = 'CENTER'
    #I.topMargin = 5  
    return I
    
def lookupalarm(alarm):
    img1 = None
    img2 = None
    if alarm[0] == 'A':
        img1 = "bell.png"
    elif alarm[0] == 'M':
        img1 = "notification.png"
    if len(alarm) > 5 and alarm[5] == 'M':
        img2 = "notification.png"
    return (img1, img2)
    
def combinecolumns(prm1, prm2, alarm):
    processed = "<font name=" + calfont + "Bold textColor=red>" + prm1 + "</font>" + "   " + "<font name=" + calfont + "Bold textColor=blue>" + prm2 + "</font>"
    if len(alarm) > 0:
        (alarmimg1, alarmimg2) = lookupalarm(alarm)
        if alarmimg1 is not None:
            inlineimg1 = "<img src=" + alarmimg1 + " width='10' height='10' valign='-2'/>"
        else:
            inlineimg1 = ""
        if alarmimg2 is not None:
            inlineimg2 = "<img src=" + alarmimg2 + " width='10' height='10' valign='-2'/>"
        else:
            inlineimg2 = ""
        processed = processed + "   " + inlineimg1 + inlineimg2
    paragraph = Paragraph(processed, matrixtimlocStyle )
    return paragraph
    
def processdescription(textpar):
    calimage = None
    dtimgeventpos = textpar.find("n[i")
    if dtimgeventpos >= 0:
        imgcode = textpar[dtimgeventpos+3:dtimgeventpos+6]
        processed = textpar[:dtimgeventpos] + textpar[dtimgeventpos+7:]
        calimage = lookupimage(imgcode)
        paragraph = Paragraph(processed, matrixdesStyle )
    else:
        paragraph = Paragraph(textpar, matrixdesStyle )
    return (paragraph, calimage)
  
def processheader(textpar):
    dtheaeventpos = textpar.find("<h")
    if dtheaeventpos == 0:
        closingtagpos = textpar.find("</h")
        if textpar[dtheaeventpos+2] == '3':
            mask = textpar[4:closingtagpos]
            processed = Paragraph(mask, style = matrixsumheadingStyle)
    elif dtheaeventpos > 0:
        mask = splicedheader(textpar, dtheaeventpos)
        processed = Paragraph(mask, style = matrixsumStyle)
    else:
        processed = Paragraph(textpar, style = matrixsumStyle)
    return processed
    
def splicedheader(textpar, index):
    closingtagpos = textpar.find("</h")
    part1 = textpar[:index]
    part2 = textpar[index+4:closingtagpos]
    processed = "<font name=" + calfont + "Bold size=10>" + part1 + "</font>" + "<font name=" + calfont + "Bold size=12>" + part2 + "</font>"
    return processed

def fillWeekReports(first_week, countdays):
    weekreps = []
    countweekreps = math.ceil((last_week - first_week + 1) / 2)
    for i in range(countweekreps):
        weekreps.append(WeekReport())
    d = str(year) + "-W" + str(first_week)
    datecal = datetime.strptime(d + '-1', '%G-W%V-%u')
    cal_day = str(int(str(datecal)[8:10]))
    wk = 0
    for i in range(countweekreps):
        weekreportname = "Familienet" + str(i) + ".pdf"
        doc = SimpleDocTemplate(weekreportname, pagesize=landscape(A4))
        storypdf=[]
        for j in range(7):
            weekreps[i].append0_Header(j, weekdaynames[j] + " " + cal_day, headerStyle)
            datecal += timedelta(days=1)
            cal_day = str(int(str(datecal)[8:10]))
        for j in range(len(monthevents)):
            if wk == monthevents[j].weeknr:
                weekreps[i].append0_Paragraph(monthevents[j].weekday, monthevents[j].summary, weeksumStyle)
                weekreps[i].append0_Paragraph(monthevents[j].weekday, monthevents[j].starttime + "-" + monthevents[j].endtime, weektimStyle)
                weekreps[i].append0_Paragraph(monthevents[j].weekday, monthevents[j].location, weeklocStyle)
                weekreps[i].append0_Paragraph(monthevents[j].weekday, monthevents[j].description, weekdesStyle)
        wk += 1
        if wk < last_week - first_week + 1:
            for j in range(7):
                weekreps[i].append1_Header(j, weekdaynames[j] + " " + cal_day, headerStyle)
                datecal += timedelta(days=1)
                cal_day = str(int(str(datecal)[8:10]))
            for j in range(len(monthevents)):
                if wk == monthevents[j].weeknr:
                    weekreps[i].append1_Paragraph(monthevents[j].weekday, monthevents[j].summary, weeksumStyle)
                    weekreps[i].append1_Paragraph(monthevents[j].weekday, monthevents[j].starttime + "-" + monthevents[j].endtime, weektimStyle)
                    weekreps[i].append1_Paragraph(monthevents[j].weekday, monthevents[j].location, weeklocStyle)
                    weekreps[i].append1_Paragraph(monthevents[j].weekday, monthevents[j].description, weekdesStyle)
        tbl_data = [
        [weekreps[i].h0[0], weekreps[i].h0[1], weekreps[i].h0[2], weekreps[i].h0[3], weekreps[i].h0[4], weekreps[i].h0[5], weekreps[i].h0[6]],
        [weekreps[i].p0[0], weekreps[i].p0[1], weekreps[i].p0[2], weekreps[i].p0[3], weekreps[i].p0[4], weekreps[i].p0[5], weekreps[i].p0[6]],
        [weekreps[i].h1[0], weekreps[i].h1[1], weekreps[i].h1[2], weekreps[i].h1[3], weekreps[i].h1[4], weekreps[i].h1[5], weekreps[i].h1[6]],
        [weekreps[i].p1[0], weekreps[i].p1[1], weekreps[i].p1[2], weekreps[i].p1[3], weekreps[i].p1[4], weekreps[i].p1[5], weekreps[i].p1[6]]
        ]
        tbl = Table(tbl_data, repeatRows=0, colWidths=[1.6*inch])
        tbl.setStyle(weekStyle)
        storypdf.append(title)
        storypdf.append(tbl)
        doc.build(storypdf)
        weekreps[i].clear()
        wk += 1
    return
    
def fillMatrixReports(countdays):
    matrixreps = []
    countmatrixReports = math.ceil(countdays / (rowsmatrixreport * columsmatrixreport))
    for i in range(countmatrixReports):
        matrixreps.append(MatrixReport())
    matrixdayhea =  [[] for _ in range(500)] 
    matrixdayheaindex = 0
    matrixdaypar =  [[] for _ in range(500)] 
    matrixdayparindex = 0
    indexreports = 0
    calimage = None
    for i in range(rowsmatrixreport):
        for j in range(columsmatrixreport):
            matrixreps[indexreports].h[i][j] = []
            matrixreps[indexreports].p[i][j] = []
    col = 0
    row = 0
    eventday = -1
    headerplaced = False
    matrixreportname = "Familienet" + str(indexreports) + ".pdf"
    doc = SimpleDocTemplate(matrixreportname, pagesize=landscape(A4))
    storypdf=[]
    for indexevents in range(len(monthevents)):
        if eventday == -1:
             eventday = monthevents[indexevents].dayyear
        if eventday != monthevents[indexevents].dayyear:
            matrixreps[indexreports].h[row][col].append(matrixdayhea[matrixdayheaindex])
            matrixreps[indexreports].p[row][col].append(matrixdaypar[matrixdayparindex])
            matrixdayheaindex += 1
            matrixdayparindex += 1
            col += 1
            eventday = monthevents[indexevents].dayyear
            headerplaced = False
            if col == columsmatrixreport:
                col = 0
                row += 1
                if row == rowsmatrixreport:
                    row = 0
                    tbl_data = [
    [matrixreps[indexreports].h[0][0], matrixreps[indexreports].h[0][1], matrixreps[indexreports].h[0][2]],
    [matrixreps[indexreports].p[0][0], matrixreps[indexreports].p[0][1], matrixreps[indexreports].p[0][2]],
    [matrixreps[indexreports].h[1][0], matrixreps[indexreports].h[1][1], matrixreps[indexreports].h[1][2]],
    [matrixreps[indexreports].p[1][0], matrixreps[indexreports].p[1][1], matrixreps[indexreports].p[1][2]],
    [matrixreps[indexreports].h[2][0], matrixreps[indexreports].h[2][1], matrixreps[indexreports].h[2][2]],   
    [matrixreps[indexreports].p[2][0], matrixreps[indexreports].p[2][1], matrixreps[indexreports].p[2][2]],
    [matrixreps[indexreports].h[3][0], matrixreps[indexreports].h[3][1], matrixreps[indexreports].h[3][2]],   
    [matrixreps[indexreports].p[3][0], matrixreps[indexreports].p[3][1], matrixreps[indexreports].p[3][2]]
                    ]
                    tbl = Table(tbl_data, repeatRows=0, rowHeights=None, colWidths=[3.75*inch])
                    tbl.setStyle(matrixStyle)
                    storypdf.append(tbl)
                    doc.build(storypdf)
                    matrixreps[indexreports].clear()
                    indexreports += 1
                    matrixreportname = "Familienet" + str(indexreports) + ".pdf"
                    doc = SimpleDocTemplate(matrixreportname, pagesize=landscape(A4))
                    storypdf=[]
        if not headerplaced:
            headerpar = Paragraph(weekdaynames[monthevents[indexevents].weekday] + " " + str(monthevents[indexevents].day) + " " + monthnames[monthevents[indexevents].month-1], headerStyle)
            matrixdayhea[matrixdayheaindex].append(headerpar)
            headerplaced = True
        paragraph = processheader(monthevents[indexevents].summary)
        matrixdaypar[matrixdayparindex].append(paragraph)
        paragraph = combinecolumns(monthevents[indexevents].starttime + "-" + monthevents[indexevents].endtime,  monthevents[indexevents].location, monthevents[indexevents].alarm)
        matrixdaypar[matrixdayparindex].append(paragraph)
        (paragraph, calimage) = processdescription(monthevents[indexevents].description)
        matrixdaypar[matrixdayparindex].append(paragraph)
        if calimage is not None:
            matrixdaypar[matrixdayparindex].append(Spacer(width=10, height=10))
            matrixdaypar[matrixdayparindex].append(calimage)        
    matrixreps[indexreports].h[row][col].append(matrixdayhea[matrixdayheaindex])
    matrixreps[indexreports].p[row][col].append(matrixdaypar[matrixdayparindex])
    tbl_data = [
    [matrixreps[indexreports].h[0][0], matrixreps[indexreports].h[0][1], matrixreps[indexreports].h[0][2]],
    [matrixreps[indexreports].p[0][0], matrixreps[indexreports].p[0][1], matrixreps[indexreports].p[0][2]],
    [matrixreps[indexreports].h[1][0], matrixreps[indexreports].h[1][1], matrixreps[indexreports].h[1][2]],
    [matrixreps[indexreports].p[1][0], matrixreps[indexreports].p[1][1], matrixreps[indexreports].p[1][2]],
    [matrixreps[indexreports].h[2][0], matrixreps[indexreports].h[2][1], matrixreps[indexreports].h[2][2]],   
    [matrixreps[indexreports].p[2][0], matrixreps[indexreports].p[2][1], matrixreps[indexreports].p[2][2]],
    [matrixreps[indexreports].h[3][0], matrixreps[indexreports].h[3][1], matrixreps[indexreports].h[3][2]],   
    [matrixreps[indexreports].p[3][0], matrixreps[indexreports].p[3][1], matrixreps[indexreports].p[3][2]]
    ]
    tbl = Table(tbl_data, repeatRows=0, rowHeights=None, colWidths=[3.75*inch])
    tbl.setStyle(matrixStyle)
    storypdf.append(tbl)
    doc.build(storypdf)
    matrixreps[indexreports].clear()
    return
    
if sys.platform[0] == 'l':
    path = '/home/jan/git/Familienet/Calendar'
if sys.platform[0] == 'w':
    path = "C:/Users/janbo/OneDrive/Documents/GitHub/Familienet/Calendar"
os.chdir(path)
eventcal = "Familienet.ics"
in_file = open(os.path.join(path, eventcal), 'r')
count = 0
lastpos = 0
alleventslines = []
for line in in_file:
    newlinepos = line.find("\t\n")
    lastsubstring = line[lastpos:newlinepos]
    alleventslines.append(lastsubstring)
    count += 1
in_file.close()
countlines = len(alleventslines)
monthevents = []
first_week = 100
last_week = -1
splitfirstw = False
splitlastw = False
countdays = 0
last_day = -1
alarm = ""
for i in range(countlines):
    neweventpos = alleventslines[i].find("BEGIN:VEVENT")
    summaryeventpos = alleventslines[i].find("SUMMARY")
    descriptioneventpos = alleventslines[i].find("DESCRIPTION")
    locationeventpos = alleventslines[i].find("LOCATION")
    dtstarteventpos = alleventslines[i].find("DTSTART")
    dtendeventpos = alleventslines[i].find("DTEND")
    endeventpos = alleventslines[i].find("END:VEVENT")
    alarmA1pos = alleventslines[i].find("A[15]")
    alarmA2pos = alleventslines[i].find("A[60]")
    alarmM1pos = alleventslines[i].find("M[14]")
    alarmM2pos = alleventslines[i].find("M[10]")
    datevaluepos = -1
    if neweventpos == 0:
        found = 0
        alldayevent = False
    if dtstarteventpos == 0:
        eventdtstartstr = alleventslines[i][8:]
        datevaluepos = alleventslines[i].find("VALUE=DATE:")
        if datevaluepos == 8:
            eventdtstartstr = alleventslines[i][19:]
            alldayevent = True
        year = int(eventdtstartstr[:4])
        month = int(eventdtstartstr[4:6])
        day = int(eventdtstartstr[6:8])
        if alldayevent:
            starttime = ""
        else:
            starttime = eventdtstartstr[9:11] + ':' + eventdtstartstr[11:13]
        weekday = weekDay(year, month, day)
        daysinmonth = leapMonth(year, month)
        weeknr = date(year=year, month=month, day=day).isocalendar()[1]
        eventday = datetime(year, month, day)
        dayyear = eventday.timetuple().tm_yday
        if dayyear != last_day:
            countdays += 1
            last_day = dayyear
        if weeknr < first_week:
            first_week = weeknr
        if weeknr > last_week:
            last_week = weeknr
        if day < 8 and weekDay(year, month, 1) > 1:
            splitfirstw = True
        if day > daysinmonth - 7 and weekDay(year, month, daysinmonth) < 6:
            splitlastw = True
        found += 1
    if dtendeventpos == 0:
        if alldayevent:
            endtime = ""
        else:
            eventdtendstr = alleventslines[i][6:]
            endtime = eventdtendstr[9:11] + ':' + eventdtendstr[11:13]
        found += 1
    if summaryeventpos == 0:
        eventsummary = alleventslines[i][8:]
        found += 1
    if descriptioneventpos == 0:
        eventdescription = alleventslines[i][12:]
        found += 1
    if locationeventpos == 0:
        eventlocation = alleventslines[i][9:]
        alarm = ""
        found += 1
    if alarmA1pos == 0:
        alarm = alarm + "A[15]"
    if alarmA2pos == 0:
        alarm = alarm + "A[60]"
    if alarmM1pos == 0:
        alarm = alarm + "M[14]"
    if alarmM2pos == 0:
        alarm = alarm + "M[10]"
    if endeventpos == 0:
        if found == 5:
            monthevents.append(FamilienetEvent(eventdescription, eventsummary, weekday - 1, weeknr - first_week, day, eventlocation, starttime, endtime, dayyear, month, alarm))
        alarm = ""
print("Count events", len(monthevents))
pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))
pdfmetrics.registerFont(TTFont('ArialItalic', 'Arial_Italic.ttf'))
pdfmetrics.registerFont(TTFont('ArialBold', 'Arial_Bold.ttf'))
pdfmetrics.registerFont(TTFont('Trebuchet', 'Trebuchet_MS.ttf'))
pdfmetrics.registerFont(TTFont('TrebuchetlItalic', 'Trebuchet_MS_Italic.ttf'))
pdfmetrics.registerFont(TTFont('TrebuchetBold', 'Trebuchet_MS_Bold.ttf'))
pdfmetrics.registerFont(TTFont('Verdana', 'Verdana.ttf'))
pdfmetrics.registerFont(TTFont('VerdanalItalic', 'Verdana_Italic.ttf'))
pdfmetrics.registerFont(TTFont('VerdanaBold', 'Verdana_Bold.ttf'))
pdfmetrics.registerFont(TTFont('Georgia', 'Georgia.ttf'))
pdfmetrics.registerFont(TTFont('GeorgiaItalic', 'Georgia_Italic.ttf'))
pdfmetrics.registerFont(TTFont('GeorgiaBold', 'Georgia_Bold.ttf'))
pdfmetrics.registerFont(TTFont('TimesNewRoman', 'Times_New_Roman.ttf'))
pdfmetrics.registerFont(TTFont('TimesNewRomanItalic', 'Times_New_Roman_Italic.ttf'))
pdfmetrics.registerFont(TTFont('TimesNewRomanBold', 'Times_New_Roman_Bold.ttf'))
pdfmetrics.registerFont(TTFont('CourierNew', 'Courier_New.ttf'))
pdfmetrics.registerFont(TTFont('CourierNewItalic', 'Courier_New_Italic.ttf'))
pdfmetrics.registerFont(TTFont('CourierNewBold', 'Courier_New_Bold.ttf'))
fillWeekReports(first_week, countdays)
#fillMatrixReports(countdays)
merger = PdfWriter()
for i in range(10):
    if os.path.isfile("Familienet" + str(i) + ".pdf"):
        inputpdf = open("Familienet" + str(i) + ".pdf", "rb")
        merger.append(inputpdf)
    else:
        break
output = open("Totaal.pdf", "wb")
merger.write(output)
merger.close()
output.close()
