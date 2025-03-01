import os
import sys
import math
import unicodedata
from pathlib import Path
from datetime import datetime, date, timedelta
from ics import Calendar, Event
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.pagesizes import LETTER, A4, landscape, portrait
from reportlab.lib.units import inch
from reportlab.lib.colors import blue, green, black, red, pink, gray, brown, purple, orange, yellow, white
from reportlab.pdfbase import pdfmetrics  
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import registerFontFamily
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, Image, Spacer, Frame
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER

startdate = datetime(1990,1,1)
datecal = datetime.now()
calfont = "LiberationSerif"
weekreps = []
columsmatrixreport = 3
rowsmatrixreport = 4
columssquarereport = 4
rowssquarereport = 4
rowscolumnreport = 15
styles = getSampleStyleSheet()
#styles.list()
sumfontsize = [[] for _ in range(300)]
sumfontsize[ord('c')] = [12, 14]
sumfontsize[ord('m')] = [11, 13]
sumfontsize[ord('w')] = [11, 13]
sumfontsize[ord('s')] = [11, 13]
#=================================================================================================================================
titleStyle = ParagraphStyle('tit', parent=styles['Normal'], fontName = calfont, fontSize = 13, textColor = black, alignment=TA_CENTER, leading = 14, spaceAfter = 3)
#=================================================================================================================================
cheaderStyle = ParagraphStyle('chea', parent=styles['Normal'], fontName = calfont, fontSize = 12, spaceAfter = 2, textColor = orange, alignment=TA_CENTER, underlineOffset = -3, underlineWidth = 0.5, underlineColor = gray, backColor = "#DAEEC9", leftIndent = 75, rightIndent = 75, leading = 14)
cheaderwkeStyle = ParagraphStyle('cheawke', parent=styles['Normal'], fontName = calfont, fontSize = 12, spaceAfter = 2, textColor = orange, alignment=TA_CENTER, underlineOffset = -3, underlineWidth = 0.5, underlineColor = gray, backColor = "#FEDDB9", leftIndent = 75, rightIndent = 75, leading = 14)
wheaderStyle = ParagraphStyle('whea', parent=styles['Normal'], fontName = calfont, fontSize = 12, textColor = orange, alignment=TA_LEFT, backColor = "#DAEEC9", leading = 14)
wheaderwkeStyle = ParagraphStyle('wheawke', parent=styles['Normal'], fontName = calfont, fontSize = 12, textColor = orange, alignment=TA_LEFT, backColor = "#FEDDB9", leading = 14)
mheaderStyle = ParagraphStyle('mhea', parent=styles['Normal'], fontName = calfont, fontSize = 12, textColor = brown, alignment=TA_CENTER, leading = 14, underlineOffset = -3, underlineWidth = 0.5, underlineColor = gray, borderWidth = 0, borderColor = "#000000", backColor = "#DAEEC9", leftIndent = 24, rightIndent = 24)
mheaderwkeStyle = ParagraphStyle('mheawke', parent=styles['Normal'], fontName = calfont, fontSize = 12, textColor = brown, alignment=TA_CENTER, leading = 15, underlineOffset = -3, underlineWidth = 0.5, underlineColor = gray, borderWidth = 0, borderColor = "#000000", backColor = "#FEDDB9", leftIndent = 24, rightIndent = 24)
sheaderStyle = ParagraphStyle('shea', parent=styles['Normal'], fontName = calfont, fontSize = 12, textColor = brown, alignment=TA_CENTER, leading = 15, underlineOffset = -3, underlineWidth = 0.5, underlineColor = gray, borderWidth = 0, borderColor = "#000000", backColor = "#DAEEC9", leftIndent = 24, rightIndent = 24)
sheaderwkeStyle = ParagraphStyle('sheawke', parent=styles['Normal'], fontName = calfont, fontSize = 12, textColor = brown, alignment=TA_CENTER, leading = 15, underlineOffset = -3, underlineWidth = 0.5, underlineColor = gray, borderWidth = 0, borderColor = "#000000", backColor = "#FEDDB9", leftIndent = 24, rightIndent = 24)
#=================================================================================================================================
weeksumStyle = ParagraphStyle('wsum', parent=styles['Normal'], fontName = calfont + "Bold", fontSize = sumfontsize[ord('w')][0], textColor = green, leading = 14)
weeklocStyle = ParagraphStyle('wloc', parent=styles['Normal'], fontName = calfont + "Italic", fontSize = 9, textColor = blue, leading = 10)
weekdesStyle = ParagraphStyle('wdes', parent=styles['Normal'], fontName = calfont, fontSize = 9, textColor = purple, leading = 10, hyphenationLang="nl_NL")
weektimStyle = ParagraphStyle('wtim', parent=styles['Normal'], fontName = calfont + "BoldItalic", fontSize = 9, textColor = red, leading = 10)
#=================================================================================================================================
matrixsumStyle = ParagraphStyle('msum', parent=styles['Normal'], fontName = calfont + "Bold", fontSize = sumfontsize[ord('m')][0], textColor = green, alignment=TA_CENTER, leading = 15)
matrixdesStyle = ParagraphStyle('mdes', parent=styles['Normal'], fontName = calfont, fontSize = 10, textColor = purple, alignment=TA_CENTER, leading = 11, hyphenationLang="nl_NL")
matrixtimlocStyle = ParagraphStyle('mloctim', parent=styles['Normal'], fontName = calfont, fontSize = 9, textColor = red, alignment=TA_CENTER, leading = 10)
#=================================================================================================================================
squaresumStyle = ParagraphStyle('ssum', parent=styles['Normal'], fontName = calfont + "Bold", fontSize = sumfontsize[ord('s')][0], textColor = green, alignment=TA_CENTER, leading = 14)
squaredesStyle = ParagraphStyle('sdes', parent=styles['Normal'], fontName = calfont, fontSize = 10, textColor = purple, alignment=TA_CENTER, leading = 11, hyphenationLang="nl_NL")
squaretimlocStyle = ParagraphStyle('sloctim', parent=styles['Normal'], fontName = calfont, fontSize = 9, textColor = red, alignment=TA_CENTER, leading = 10)
#=================================================================================================================================
columnsumStyle = ParagraphStyle('csum', parent=styles['Normal'], fontName = calfont + "Bold", fontSize = sumfontsize[ord('c')][0], textColor = green, alignment=TA_CENTER, leading = 15)
columndesStyle = ParagraphStyle('cdes', parent=styles['Normal'], fontName = calfont, fontSize = 10, textColor = purple, alignment=TA_CENTER, leading = 11)
columntimlocStyle = ParagraphStyle('ctimloc', parent=styles['Normal'], fontName = calfont, fontSize = 8, textColor = red, alignment=TA_CENTER, leading = 9)
#================================================================================================================================
version = "Maart 2025"
weekdaynames = ["Maandag","Dinsdag","Woensdag","Donderdag","Vrijdag","Zaterdag","Zondag"]
monthnames = ["Januari","Februari","Maart","April","Mei","Juni","Juli","Augustus", "September","Oktober","November","December"]
weekStyle = [
('FONTSIZE', (0, 1), (-1, 1), 10),
('VALIGN',(0,0),(-1,-1),'TOP'),
('ALIGN',(0,0),(3,1),'CENTER')
]
matrixStyle = [
('VALIGN',(0,0),(-1,-1),'TOP'),
('ALIGN',(0,0),(3,1),'CENTER')
]
squareStyle = [
('VALIGN',(0,0),(-1,-1),'TOP'),
('ALIGN',(0,0),(3,1),'CENTER')
]

class ColumnReport:
    d = []
    
    def append_Paragraph(self, paragraph, style):
        textpar = Paragraph(paragraph, style)
        self.d.append(textpar)
    
    def clear(self):
        while len(self.d) > 0:
            self.d.pop()
    
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

    def tabledata(self):
        return [[self.h0[0], self.h0[1], self.h0[2], self.h0[3], self.h0[4], self.h0[5], self.h0[6]],
                [self.p0[0], self.p0[1], self.p0[2], self.p0[3], self.p0[4], self.p0[5], self.p0[6]],
                [self.h1[0], self.h1[1], self.h1[2], self.h1[3], self.h1[4], self.h1[5], self.h1[6]],
                [self.p1[0], self.p1[1], self.p1[2], self.p1[3], self.p1[4], self.p1[5], self.p1[6]]]

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

    def tabledata(self):
        return [[self.h[0][0], self.h[0][1], self.h[0][2]],
                [self.p[0][0], self.p[0][1], self.p[0][2]],
                [self.h[1][0], self.h[1][1], self.h[1][2]],
                [self.p[1][0], self.p[1][1], self.p[1][2]],
                [self.h[2][0], self.h[2][1], self.h[2][2]],   
                [self.p[2][0], self.p[2][1], self.p[2][2]],
                [self.h[3][0], self.h[3][1], self.h[3][2]],   
                [self.p[3][0], self.p[3][1], self.p[3][2]]]

class SquareReport:
    h = [[0 for i in range(columssquarereport)] for j in range(rowssquarereport)] 
    p = [[0 for i in range(columssquarereport)] for j in range(rowssquarereport)] 
 
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

    def tabledata(self):
        return [[self.h[0][0], self.h[0][1], self.h[0][2], self.h[0][3]],
                [self.p[0][0], self.p[0][1], self.p[0][2], self.p[0][3]],
                [self.h[1][0], self.h[1][1], self.h[1][2], self.h[1][3]],
                [self.p[1][0], self.p[1][1], self.p[1][2], self.p[1][3]],
                [self.h[2][0], self.h[2][1], self.h[2][2], self.h[2][3]],  
                [self.p[2][0], self.p[2][1], self.p[2][2], self.p[2][3]],
                [self.h[3][0], self.h[3][1], self.h[3][2], self.h[3][3]],  
                [self.p[3][0], self.p[3][1], self.p[3][2], self.p[3][3]]]
    
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
        
def find_all_occurrences(textpar, sub):
    index_of_occurrences = []
    current_index = 0
    while True:
        current_index = textpar.find(sub, current_index)
        if current_index == -1:
            return index_of_occurrences
        else:
            index_of_occurrences.append(current_index)
            current_index += len(sub)
        
def processreport(t):
    merger = PdfWriter()
    for i in range(50):
        if os.path.isfile("PDF/Familienet" + str(i) + ".pdf"):
            inputpdf = open("PDF/Familienet" + str(i) + ".pdf", "rb")
            merger.append(inputpdf)
            inputpdf.close()
        else:
            break
    output = open("PDF/" + t + version + ".pdf", "wb")
    merger.write(output)
    merger.close()
    output.close()
    for i in range(50):
        if os.path.isfile("PDF/Familienet" + str(i) + ".pdf"):
            os.remove("PDF/Familienet" + str(i) + ".pdf")
        
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

def lookupflag(imgcode):
    flagimg = "Flags/nl.png"
    if imgcode == "HOL":
        flagimg = "Flags/nl.png"
    if imgcode == "POL":
        flagimg = "Flags/pl.png"
    if imgcode == "FRA":
        flagimg = "Flags/fr.png"
    if imgcode == "AUT":
        flagimg = "Flags/at.png"
    if imgcode == "ESP":
        flagimg = "Flags/es.png"
    if imgcode == "DEU":
        flagimg = "Flags/de.png"
    return flagimg
    
def lookupimage(imgcode):
    newcalendar = [[] for _ in range(300)]
    newcalendar[139] = "Photos/ASSO.jpg"
    newcalendar[154] = "Photos/DoReMichiel.jpg"
    newcalendar[156] = "Photos/VivaLaFrance.jpg"
    newcalendar[157] = "Photos/Herfst2.jpg"
    newcalendar[107] = "Photos/kerstversiering.png"
    newcalendar[164] = "Photos/sinterklaas.jpg"
    newcalendar[171] = "Photos/MiniPrideWalk12juni-2.jpg"
    newcalendar[182] = "Photos/Herfst3.jpg"
    newcalendar[183] = "Photos/Herfst1.jpg"
    newcalendar[186] = "Photos/JoAnn.jpg"
    newcalendar[191] = "Photos/6617f09c1049a.jpg"
    newcalendar[190] = "Photos/keukenhof-lisse-molen.jpg"
    newcalendar[204] = "Photos/R.jpeg"
    newcalendar[214] = "Photos/draaiorgelpineutje.jpg"
    newcalendar[262] = "Photos/spain_PNG58.png"
    newcalendar[266] = "Photos/AlpenWeide.jpeg"
    newcalendar[269] = "Photos/Los-del-Sol-cantando-mariachi.jpg"
    newcalendar[271] = "Photos/Conny-en-Ronald.png"
    img = newcalendar[int(imgcode)]
    I = Image(img)
    return I

def lookupalarm(alarm):
    img1 = None
    img2 = None
    if alarm[0] == 'A':
        img1 = "Alarms/bell.png"
    elif alarm[0] == 'M':
        img1 = "Alarms/notification.png"
    if len(alarm) > 5 and alarm[5] == 'M':
        img2 = "Alarms/notification.png"
    return (img1, img2)

def lookupemoji(imgcode):
    emojitable = [[] for _ in range(3000)]
    emojitable[884] = "Emojis/374.png" #uiteten
    emojitable[900] = "Emojis/384.png" #kerstboom
    emojitable[905] = "Emojis/389.png" #bingo
    emojitable[932] = "Emojis/3A4.png" #microfoon
    emojitable[933] = "Emojis/3A5.png" #filmcamera
    emojitable[948] = "Emojis/3B4.png" #sjoelen
    emojitable[953] = "Emojis/3B9.png" #klavier
    emojitable[954] = "Emojis/3BA.png" #trompet
    emojitable[955] = "Emojis/3BB.png" #viool
    emojitable[956] = "Emojis/3BC.png" #muzieksleutel
    emojitable[1108] = "Emojis/454.png" #overhemd
    emojitable[1155] = "Emojis/483.png" #danseres
    emojitable[1168] = "Emojis/490.png" #bloemschikken
    emojitable[1171] = "Emojis/493.png" #valentijn
    emojitable[1183] = "Emojis/49F.png" #wenskaart
    emojitable[1216] = "Emojis/4C0.png" #goudenplaat
    emojitable[1242] = "Emojis/4DA.png" #bibliotheek
    emojitable[1543] = "Emojis/607.png" #halo
    emojitable[1550] = "Emojis/60E.png" #zonnebril
    emojitable[1718] = "Emojis/6B6.png" #wandelen
    emojitable[1744] = "Emojis/6D0.png" #kerkdienst
    emojitable[1602] = "Emojis/642.png" #glimlach
    emojitable[2360] = "Emojis/938.png" #bewegenopmuziek
    emojitable[2366] = "Emojis/93E.png" #sportenspel
    emojitable[2503] = "Emojis/9C7.png" #wafel
    emojitable[2711] = "Emojis/A97.png" #accordeon
    emojiint = int(imgcode, 16)
    emojiimg = emojitable[emojiint]
    return emojiimg

def combinecolumns(prm1, prm2, alarm):
    processed = "<font name=" + calfont + " textColor=red>" + prm1 + "</font>" + "   " + "<font name=" + calfont + " textColor=blue>" + prm2 + "</font>"
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
    return processed
    
def processimage(countevents, imgcode):
    index = 0
    calwimage = None
    if len(imgcode) > 0:
        minevents = 100
        for d in range(7):
            if countevents[d] < minevents:
                minevents = countevents[d]
                index = d 
        calwimage = lookupimage(imgcode)
        ratio = calwimage.drawWidth / calwimage.drawHeight
        calwimage.drawHeight = 0.8 * inch
        calwimage.drawWidth = ratio * calwimage.drawHeight
        calwimage.hAlign = 'CENTER'
    return (index, calwimage)
    
def processwtime(textpar, alarm):
    processed = textpar
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
    return processed
    
def processwdescription(textpar):
    imgcode = ""
    dtimgeventpos = textpar.find("[i")
    if dtimgeventpos >= 0:
        imgcode = textpar[dtimgeventpos+2:dtimgeventpos+5]
    return imgcode
    
def processcdescription(textpar, s):
    calimage = None
    dtimgeventpos = textpar.find("[i")
    if dtimgeventpos >= 0:
        imgcode = textpar[dtimgeventpos+2:dtimgeventpos+5]
        processed = textpar[:dtimgeventpos-1] + textpar[dtimgeventpos+6:]
        calimage = lookupimage(imgcode)
        paragraph = Paragraph(processed, s)
    else:
        paragraph = Paragraph(textpar, s)
    return (paragraph, calimage)

def processmsdescription(textpar, s):
    imgcode = ""
    imgpos = ''
    dtimgeventpos = textpar.find("[i")
    if dtimgeventpos >= 0:
        imgcode = textpar[dtimgeventpos+2:dtimgeventpos+5]
        imgpos = textpar[dtimgeventpos-1]
        processed = textpar[:dtimgeventpos-1] + textpar[dtimgeventpos+6:]
        paragraph = Paragraph(processed, s)
    else:
        paragraph = Paragraph(textpar, s)
    return (paragraph, imgcode, imgpos)

def processsummary(textpar, t):
    processed = textpar
    tagpos = textpar.find("<h")
    if tagpos >= 0:
        closingtagpos = textpar.find("</h")
        part1 = textpar[:tagpos]
        part2 = textpar[tagpos+4:closingtagpos]
        processed = "<font name=" + calfont + "Bold size=" + str(sumfontsize[ord(t)][0]) + ">" + part1 + "</font>" + "<font name=" + calfont + "Bold size=" + str(sumfontsize[ord(t)][1]) + ">" + part2 + "</font>" + textpar[closingtagpos+5:]
    fls = find_all_occurrences(processed, "[f")
    if len(fls) > 0:
        for f in range(len(fls) - 1, -1, -1):
            g = fls[f]
            flagimg = lookupflag(processed[g+2:g+5]           )
            inlineimg = "<img src=" + flagimg + " width='15' height='10' valign='-2'/>"
            processed = processed.replace(processed[g:g+5], inlineimg)
    ems = find_all_occurrences(processed, "[e")
    if len(ems) > 0:
        for e in range(len(ems) - 1, -1, -1):
            h = ems[e]
            emojiimg = lookupemoji(processed[h+2:h+5])
            inlineimg = "<img src=" + emojiimg + " width='18' height='18' valign='-6'/>"
            processed = processed.replace(processed[h:h+5], inlineimg)
    return processed

def fillColumnReports(countdays):
    columnreps = []
    countcolumnReports = 4
    for i in range(countcolumnReports):
        columnreps.append(ColumnReport())
    i = 0
    columnreportname = "PDF/Familienet" + str(i) + ".pdf"
    doc = SimpleDocTemplate(columnreportname, pagesize=portrait(A4), rightMargin=5, leftMargin=5, topMargin=5, bottomMargin=5)
    storypdf=[]
    eventday = -1
    rows = 0
    openreport = False
    for indexevents in range(len(monthevents)):
        openreport = True
        if eventday == -1 or eventday != monthevents[indexevents].dayyear:
            header = "<u>" + weekdaynames[monthevents[indexevents].weekday] + " " + str(monthevents[indexevents].day) + " " + monthnames[monthevents[indexevents].month-1] + "</u>"
            if monthevents[indexevents].weekday < 5:
                columnreps[i].append_Paragraph(header, cheaderStyle)
            else:
                columnreps[i].append_Paragraph(header, cheaderwkeStyle)
            eventday = monthevents[indexevents].dayyear
        columnreps[i].d.append(Paragraph(processsummary(monthevents[indexevents].summary, 'c'), columnsumStyle))
        columnreps[i].d.append(Paragraph(combinecolumns(monthevents[indexevents].starttime + "-" + monthevents[indexevents].endtime,  monthevents[indexevents].location, monthevents[indexevents].alarm), columntimlocStyle))
        (paragraph, calimage) = processcdescription(monthevents[indexevents].description, columndesStyle)
        columnreps[i].d.append(paragraph)
        if calimage is not None:
            ratio = calimage.drawWidth / calimage.drawHeight
            calimage.drawHeight = 1.0 * inch
            calimage.drawWidth = ratio * calimage.drawHeight
            margin = (7.5 * inch - calimage.drawWidth) / 2
            sp1 = Spacer(margin - 0.3 * inch, calimage.drawHeight)
            sp2 = Spacer(margin, calimage.drawHeight)
            columnreps[i].d.append(Table([[sp1, calimage, sp2]], rowHeights = [1.1 * inch]))
        rows += 1
        if rows == rowscolumnreport:
            tbl_data = [[columnreps[i].d]]
            tbl = Table(tbl_data, repeatRows=0, colWidths=[7.5*inch])
            storypdf.append(Paragraph(version, titleStyle))
            storypdf.append(tbl)
            doc.build(storypdf)
            columnreps[i].clear()
            i += 1
            storypdf=[]
            columnreportname = "PDF/Familienet" + str(i) + ".pdf"
            doc = SimpleDocTemplate(columnreportname, pagesize=portrait(A4), rightMargin=5, leftMargin=5, topMargin=5, bottomMargin=5)
            eventday = -1
            rows = 0
            openreport = False
    if openreport:
        tbl_data = [[columnreps[i].d]]
        tbl = Table(tbl_data, repeatRows=0, colWidths=[7.5*inch])
        storypdf.append(Paragraph(version, titleStyle))
        storypdf.append(tbl)
        doc.build(storypdf)
        columnreps[i].clear()
    return
    
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
        weekreportname = "PDF/Familienet" + str(i) + ".pdf"
        doc = SimpleDocTemplate(weekreportname, pagesize=landscape(A4), rightMargin=5, leftMargin=5, topMargin=5, bottomMargin=5)
        storypdf=[]
        imgcode = ""
        imgday = -1
        countevents = [0,0,0,0,0,0,0]
        for j in range(7):
            if j < 5:
                weekreps[i].append0_Header(j, weekdaynames[j] + " " + cal_day, wheaderStyle)
            else:
                weekreps[i].append0_Header(j, weekdaynames[j] + " " + cal_day, wheaderwkeStyle)
            datecal += timedelta(days=1)
            cal_day = str(int(str(datecal)[8:10]))
        for j in range(len(monthevents)):
            if wk == monthevents[j].weeknr:
                if imgcode == "":
                    imgcode = processwdescription(monthevents[j].description)
                    if len(imgcode) > 0:
                        imgday = monthevents[j].weekday
                index = monthevents[j].weekday
                weekreps[i].p0[index].append(Paragraph(processsummary(monthevents[j].summary, 'w'), weeksumStyle))
                weekreps[i].append0_Paragraph(index, processwtime(monthevents[j].starttime + "-" + monthevents[j].endtime, monthevents[j].alarm), weektimStyle)
                weekreps[i].append0_Paragraph(index, monthevents[j].location, weeklocStyle)
                if imgday == index:
                    wimgpos = monthevents[j].description.find("[i")
                    if wimgpos >= 0:
                        monthevents[j].description = monthevents[j].description[:wimgpos-1] + monthevents[j].description[wimgpos+6:]
                weekreps[i].append0_Paragraph(index, monthevents[j].description, weekdesStyle)
                countevents[index] += 1
        (index, calwimage) = processimage(countevents, imgcode)
        if calwimage is not None:
            weekreps[i].p0[index].append(calwimage)
            imgcode = ""
            imgday = -1
        wk += 1
        if wk < last_week - first_week + 1:
            imgcode = ""
            countevents = [0,0,0,0,0,0,0]
            for j in range(7):
                if j < 5:
                    weekreps[i].append1_Header(j, weekdaynames[j] + " " + cal_day, wheaderStyle)
                else:
                    weekreps[i].append1_Header(j, weekdaynames[j] + " " + cal_day, wheaderwkeStyle) 
                datecal += timedelta(days=1)
                cal_day = str(int(str(datecal)[8:10]))
            for j in range(len(monthevents)):
                if wk == monthevents[j].weeknr:
                    if imgcode == "":
                        imgcode = processwdescription(monthevents[j].description)
                        if len(imgcode) > 0:
                            imgday = monthevents[j].weekday
                    index = monthevents[j].weekday
                    weekreps[i].p1[index].append(Paragraph(processsummary(monthevents[j].summary, 'w'), weeksumStyle))
                    weekreps[i].append1_Paragraph(index, processwtime(monthevents[j].starttime + "-" + monthevents[j].endtime, monthevents[j].alarm), weektimStyle)
                    weekreps[i].append1_Paragraph(index, monthevents[j].location, weeklocStyle)
                    if imgday == index:
                        wimgpos = monthevents[j].description.find("[i")
                        if wimgpos >= 0:
                            monthevents[j].description = monthevents[j].description[:wimgpos-1] + monthevents[j].description[wimgpos+6:]
                    weekreps[i].append1_Paragraph(index, monthevents[j].description, weekdesStyle)
                    countevents[index] += 1
            (index, calwimage) = processimage(countevents, imgcode)
            if calwimage is not None:
                weekreps[i].p1[index].append(calwimage)
                imgcode = ""
                imgday = -1
        tbl_data = weekreps[i].tabledata()
        tbl = Table(tbl_data, repeatRows=0, colWidths=[1.6*inch])
        tbl.setStyle(weekStyle)
        storypdf.append(Paragraph(version, titleStyle))
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
    indrep = 0
    dagimgcode = ""
    dagimgpos = ''
    imgcode = ""
    imgpos = ''
    for i in range(rowsmatrixreport):
        for j in range(columsmatrixreport):
            matrixreps[indrep].h[i][j] = []
            matrixreps[indrep].p[i][j] = []
    col = 0
    row = 0
    eventday = -1
    headerplaced = False
    matrixreportname = "PDF/Familienet" + str(indrep) + ".pdf"
    doc = SimpleDocTemplate(matrixreportname, pagesize=landscape(A4), rightMargin=5, leftMargin=5, topMargin=5, bottomMargin=5)
    storypdf=[]
    for indexevents in range(len(monthevents)):
        if eventday == -1:
             eventday = monthevents[indexevents].dayyear
        if eventday != monthevents[indexevents].dayyear:
            if len(dagimgcode) > 0 and dagimgpos == 'l':
                calimage = lookupimage(dagimgcode)
                ratio = calimage.drawWidth / calimage.drawHeight
                calimage.drawWidth = ratio * 1.1 * inch
                margin = (3.75 * inch - calimage.drawWidth) / 2
                calimage.drawHeight = 1.1 * inch
                calimage.hAlign = 'CENTER'
                matrixreps[indrep].p[row][col].append(Table([[None, calimage, None]], colWidths=[margin, calimage.drawWidth, margin],  rowHeights=[1.1 * inch]))
                dagimgcode = ""
                dagimgpos = ''
                calimage = None
                col += 1
                if col == columsmatrixreport:
                    col = 0
                    row += 1
                    if row == rowsmatrixreport:
                        row = 0
                        tbl_data = matrixreps[indrep].tabledata()
                        tbl = Table(tbl_data, repeatRows=0, rowHeights=None, colWidths=[2.92*inch])
                        tbl.setStyle(matrixStyle)
                        storypdf.append(Paragraph(version, titleStyle))
                        storypdf.append(tbl)
                        doc.build(storypdf)
                        matrixreps[indrep].clear()
                        indrep += 1
                        matrixreportname = "PDF/Familienet" + str(indrep) + ".pdf"
                        doc = SimpleDocTemplate(matrixreportname, pagesize=landscape(A4), rightMargin=5, leftMargin=5, topMargin=5, bottomMargin=5)
                        storypdf=[]
            matrixreps[indrep].h[row][col].append(matrixdayhea[matrixdayheaindex])
            matrixreps[indrep].p[row][col].append(matrixdaypar[matrixdayparindex])
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
                    tbl_data = matrixreps[indrep].tabledata()
                    tbl = Table(tbl_data, repeatRows=0, rowHeights=None, colWidths=[3.75*inch])
                    tbl.setStyle(matrixStyle)
                    storypdf.append(Paragraph(version, titleStyle))
                    storypdf.append(tbl)
                    doc.build(storypdf)
                    matrixreps[indrep].clear()
                    indrep += 1
                    matrixreportname = "PDF/Familienet" + str(indrep) + ".pdf"
                    doc = SimpleDocTemplate(matrixreportname, pagesize=landscape(A4), rightMargin=5, leftMargin=5, topMargin=5, bottomMargin=5)
                    storypdf=[]
        if not headerplaced:
            if monthevents[indexevents].weekday == 5 or monthevents[indexevents].weekday == 6:
                headerpar = Paragraph("<u>" + weekdaynames[monthevents[indexevents].weekday] + " " + str(monthevents[indexevents].day) + " " + monthnames[monthevents[indexevents].month-1] + "</u>", mheaderwkeStyle)
            else:
                headerpar = Paragraph("<u>" + weekdaynames[monthevents[indexevents].weekday] + " " + str(monthevents[indexevents].day) + " " + monthnames[monthevents[indexevents].month-1] + "</u>", mheaderStyle)
            matrixdayhea[matrixdayheaindex].append(headerpar)
            headerplaced = True
        matrixdaypar[matrixdayparindex].append(Paragraph(processsummary(monthevents[indexevents].summary, 'm'), matrixsumStyle))
        matrixdaypar[matrixdayparindex].append(Paragraph(combinecolumns(monthevents[indexevents].starttime + "-" + monthevents[indexevents].endtime,  monthevents[indexevents].location, monthevents[indexevents].alarm), matrixtimlocStyle))
        (paragraph, imgcode, imgpos) = processmsdescription(monthevents[indexevents].description, matrixdesStyle)
        matrixdaypar[matrixdayparindex].append(paragraph)
        if len(imgcode) > 0:
            if imgpos == 'b':
                calimage = lookupimage(imgcode)
                calimage.drawHeight = 0.95*inch
                calimage.drawWidth = 0.95*inch
                calimage.hAlign = 'CENTER'
                matrixdaypar[matrixdayparindex].append(Spacer(width=10, height=10))
                matrixdaypar[matrixdayparindex].append(Table([[None, calimage, None]], colWidths=[1.1 * inch, 1.1 * inch, 1.1 * inch],  rowHeights=[1.1 * inch]))
                imgcode = ""
                ímgpos = ''
                calimage = None
            elif imgpos == 'l':
                dagimgcode = imgcode
                dagimgpos = imgpos
    matrixreps[indrep].h[row][col].append(matrixdayhea[matrixdayheaindex])
    matrixreps[indrep].p[row][col].append(matrixdaypar[matrixdayparindex])
    tbl_data = matrixreps[indrep].tabledata()
    tbl = Table(tbl_data, repeatRows=0, rowHeights=None, colWidths=[3.75*inch])
    tbl.setStyle(matrixStyle)
    storypdf.append(Paragraph(version, titleStyle))
    storypdf.append(tbl)
    doc.build(storypdf)
    matrixreps[indrep].clear()
    return

def fillSquareReports(countdays):
    squarereps = []
    countsquareReports = math.ceil(countdays / (rowssquarereport * columssquarereport))
    for i in range(countsquareReports):
        squarereps.append(SquareReport())
    squaredayhea =  [[] for _ in range(500)] 
    squaredayheaindex = 0
    squaredaypar =  [[] for _ in range(500)] 
    squaredayparindex = 0
    indrep = 0
    dagimgcode = ""
    dagimgpos = ''
    imgcode = ""
    imgpos = ''
    for i in range(rowssquarereport):
        for j in range(columssquarereport):
            squarereps[indrep].h[i][j] = []
            squarereps[indrep].p[i][j] = []
    col = 0
    row = 0
    eventday = -1
    headerplaced = False
    squarereportname = "PDF/Familienet" + str(indrep) + ".pdf"
    doc = SimpleDocTemplate(squarereportname, pagesize=landscape(A4), rightMargin=5, leftMargin=5, topMargin=5, bottomMargin=5)
    storypdf=[]
    for indexevents in range(len(monthevents)):
        if eventday == -1:
             eventday = monthevents[indexevents].dayyear
        if eventday != monthevents[indexevents].dayyear:
            if len(dagimgcode) > 0 and dagimgpos == 'l':
                calimage = lookupimage(dagimgcode)
                ratio = calimage.drawWidth / calimage.drawHeight
                calimage.drawWidth = ratio * 1.0 * inch
                margin = (2.92 * inch - calimage.drawWidth) / 2
                calimage.drawHeight = 1.0 * inch
                calimage.hAlign = 'CENTER'
                squarereps[indrep].p[row][col].append(Table([[None, calimage, None]], colWidths=[margin, calimage.drawWidth, margin],  rowHeights=[1.0 * inch]))
                dagimgcode = ""
                dagimgpos = ''
                calimage = None
                col += 1
                if col == columssquarereport:
                    col = 0
                    row += 1
                    if row == rowssquarereport:
                        row = 0
                        tbl_data = squarereps[indrep].tabledata()
                        tbl = Table(tbl_data, repeatRows=0, rowHeights=None, colWidths=[2.92*inch])
                        tbl.setStyle(squareStyle)
                        storypdf.append(Paragraph(version, titleStyle))
                        storypdf.append(tbl)
                        doc.build(storypdf)
                        squarereps[indrep].clear()
                        indrep += 1
                        squarereportname = "PDF/Familienet" + str(indrep) + ".pdf"
                        doc = SimpleDocTemplate(squarereportname, pagesize=landscape(A4), rightMargin=5, leftMargin=5, topMargin=5, bottomMargin=5)
                        storypdf=[]
            squarereps[indrep].h[row][col].append(squaredayhea[squaredayheaindex])
            squarereps[indrep].p[row][col].append(squaredaypar[squaredayparindex])
            squaredayheaindex += 1
            squaredayparindex += 1
            col += 1
            eventday = monthevents[indexevents].dayyear
            headerplaced = False
            if col == columssquarereport:
                col = 0
                row += 1
                if row == rowssquarereport:
                    row = 0
                    tbl_data = squarereps[indrep].tabledata()
                    tbl = Table(tbl_data, repeatRows=0, rowHeights=None, colWidths=[2.92*inch])
                    tbl.setStyle(squareStyle)
                    storypdf.append(Paragraph(version, titleStyle))
                    storypdf.append(tbl)
                    doc.build(storypdf)
                    squarereps[indrep].clear()
                    indrep += 1
                    squarereportname = "PDF/Familienet" + str(indrep) + ".pdf"
                    doc = SimpleDocTemplate(squarereportname, pagesize=landscape(A4), rightMargin=5, leftMargin=5, topMargin=5, bottomMargin=5)
                    storypdf=[]
        if not headerplaced:
            if monthevents[indexevents].weekday == 5 or monthevents[indexevents].weekday == 6:
                headerpar = Paragraph("<u>" + weekdaynames[monthevents[indexevents].weekday] + " " + str(monthevents[indexevents].day) + " " + monthnames[monthevents[indexevents].month-1] + "</u>", sheaderwkeStyle)
            else:
                headerpar = Paragraph("<u>" + weekdaynames[monthevents[indexevents].weekday] + " " + str(monthevents[indexevents].day) + " " + monthnames[monthevents[indexevents].month-1] + "</u>", sheaderStyle)
            squaredayhea[squaredayheaindex].append(headerpar)
            headerplaced = True
        squaredaypar[squaredayparindex].append(Paragraph(processsummary(monthevents[indexevents].summary, 's'), squaresumStyle))
        squaredaypar[squaredayparindex].append(Paragraph(combinecolumns(monthevents[indexevents].starttime + "-" + monthevents[indexevents].endtime,  monthevents[indexevents].location, monthevents[indexevents].alarm), squaretimlocStyle))
        (paragraph, imgcode, imgpos) = processmsdescription(monthevents[indexevents].description, squaredesStyle)
        squaredaypar[squaredayparindex].append(paragraph)
        if len(imgcode) > 0:
            if imgpos == 'b':
                calimage = lookupimage(imgcode)
                calimage.drawHeight = 0.95*inch
                calimage.drawWidth = 0.95*inch
                calimage.hAlign = 'CENTER'
                squaredaypar[squaredayparindex].append(Spacer(width=10, height=10))
                squaredaypar[squaredayparindex].append(Table([[None, calimage, None]], colWidths=[0.5 * inch, 0.96 * inch, 0.5 * inch],  rowHeights=[0.95 * inch]))
                imgcode = ""
                ímgpos = ''
                calimage = None
            elif imgpos == 'l':
                dagimgcode = imgcode
                dagimgpos = imgpos
    if len(imgcode) > 0 and imgpos == 'l':
        calimage = lookupimage(imgcode)
        calimage.drawHeight = 0.95*inch
        calimage.drawWidth = 0.95*inch
        calimage.hAlign = 'CENTER'
        squarereps[indrep].p[row][col].append(Table([[None, calimage, None]], colWidths=[0.5 * inch, 0.96 * inch, 0.5 * inch],  rowHeights=[0.95 * inch]))
        imgcode = ""
        ímgpos = ''
        calimage = None
        col += 1
        if col == columssquarereport:
            col = 0
            row += 1
            if row == rowssquarereport:
                row = 0
                tbl_data = squarereps[indrep].tabledata()
                tbl = Table(tbl_data, repeatRows=0, rowHeights=None, colWidths=[2.92*inch])
                tbl.setStyle(squareStyle)
                storypdf.append(Paragraph(version, titleStyle))
                storypdf.append(tbl)
                doc.build(storypdf)
                squarereps[indrep].clear()
                indrep += 1
                squarereportname = "PDF/Familienet" + str(indrep) + ".pdf"
                doc = SimpleDocTemplate(squarereportname, pagesize=landscape(A4), rightMargin=5, leftMargin=5, topMargin=5, bottomMargin=5)
                storypdf=[]
    squarereps[indrep].h[row][col].append(squaredayhea[squaredayheaindex])
    squarereps[indrep].p[row][col].append(squaredaypar[squaredayparindex])
    tbl_data = squarereps[indrep].tabledata()
    tbl = Table(tbl_data, repeatRows=0, rowHeights=None, colWidths=[2.92*inch])
    tbl.setStyle(squareStyle)
    storypdf.append(Paragraph(version, titleStyle))
    storypdf.append(tbl)
    doc.build(storypdf)
    squarereps[indrep].clear()
    return

if sys.platform[0] == 'l':
    path = '/home/jan/git/Familienet'
if sys.platform[0] == 'w':
    path = "C:/Users/janbo/OneDrive/Documents/GitHub/Familienet"
os.chdir(path)
eventcal = "Calendar/Familienet.ics"
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
pdfmetrics.registerFont(TTFont('ArialBold', 'Arial_Bold.ttf'))
pdfmetrics.registerFont(TTFont('ArialItalic', 'Arial_Italic.ttf'))
pdfmetrics.registerFont(TTFont('ArialBoldItalic', 'Arial_Bold_Italic.ttf'))
pdfmetrics.registerFont(TTFont('Verdana', 'Verdana.ttf'))
pdfmetrics.registerFont(TTFont('VerdanaBold', 'Verdana_Bold.ttf'))
pdfmetrics.registerFont(TTFont('VerdanaItalic', 'Verdana_Italic.ttf'))
pdfmetrics.registerFont(TTFont('VerdanaBoldItalic', 'Verdana_Bold_Italic.ttf'))
pdfmetrics.registerFont(TTFont('Georgia', 'Georgia.ttf'))
pdfmetrics.registerFont(TTFont('GeorgiaBold', 'Georgia_Bold.ttf'))
pdfmetrics.registerFont(TTFont('GeorgiaItalic', 'Georgia_Italic.ttf'))
pdfmetrics.registerFont(TTFont('GeorgiaBoldItalic', 'Georgia_Bold_Italic.ttf'))
pdfmetrics.registerFont(TTFont('TimesNewRoman', 'times.ttf'))
pdfmetrics.registerFont(TTFont('TimesNewRomanBold', 'timesbd.ttf'))
pdfmetrics.registerFont(TTFont('TimesNewRomanItalic', 'timesi.ttf'))
pdfmetrics.registerFont(TTFont('TimesNewRomanBoldItalic', 'timesbi.ttf'))
pdfmetrics.registerFont(TTFont('Trebuchet', 'Trebuchet_MS.ttf'))
pdfmetrics.registerFont(TTFont('TrebuchetBold', 'Trebuchet_MS_Bold.ttf'))
pdfmetrics.registerFont(TTFont('TrebuchetItalic', 'Trebuchet_MS_Italic.ttf'))
pdfmetrics.registerFont(TTFont('TrebuchetBoldItalic', 'Trebuchet_MS_Bold_Italic.ttf'))
pdfmetrics.registerFont(TTFont('Ubuntu', 'Ubuntu-Regular.ttf'))
pdfmetrics.registerFont(TTFont('UbuntuBold', 'Ubuntu-Bold.ttf'))
pdfmetrics.registerFont(TTFont('UbuntuItalic', 'Ubuntu-Italic.ttf'))
pdfmetrics.registerFont(TTFont('UbuntuBoldItalic', 'Ubuntu-BoldItalic.ttf'))
pdfmetrics.registerFont(TTFont('DejaVuSerif', 'DejaVuSerif.ttf'))
pdfmetrics.registerFont(TTFont('DejaVuSerifBold', 'DejaVuSerif-Bold.ttf'))
pdfmetrics.registerFont(TTFont('DejaVuSerifItalic', 'DejaVuSerif-Italic.ttf'))
pdfmetrics.registerFont(TTFont('DejaVuSerifBoldItalic', 'DejaVuSerif-BoldItalic.ttf'))
pdfmetrics.registerFont(TTFont('LiberationSerif', 'LiberationSerif-Regular.ttf'))
pdfmetrics.registerFont(TTFont('LiberationSerifBold', 'LiberationSerif-Bold.ttf'))
pdfmetrics.registerFont(TTFont('LiberationSerifItalic', 'LiberationSerif-Italic.ttf'))
pdfmetrics.registerFont(TTFont('LiberationSerifBoldItalic', 'LiberationSerif-BoldItalic.ttf'))
pdfmetrics.registerFont(TTFont('DancingScript', 'DancingScript-Regular.ttf'))
pdfmetrics.registerFont(TTFont('DancingScriptBold', 'DancingScript-Bold.ttf'))
pdfmetrics.registerFont(TTFont('DancingScriptItalic', 'DancingScript-Regular.ttf'))
pdfmetrics.registerFont(TTFont('DancingScriptBoldItalic', 'DancingScript-Bold.ttf'))
pdfmetrics.registerFont(TTFont('CormorantGaramond', 'CormorantGaramond-Regular.ttf'))
pdfmetrics.registerFont(TTFont('CormorantGaramondBold', 'CormorantGaramond-Bold.ttf'))
pdfmetrics.registerFont(TTFont('CormorantGaramondItalic', 'CormorantGaramond-Italic.ttf'))
pdfmetrics.registerFont(TTFont('CormorantGaramondBoldItalic', 'CormorantGaramond-BoldItalic.ttf'))
fillMatrixReports(countdays)
processreport('m')
fillColumnReports(countdays)
processreport('c')
fillSquareReports(countdays)
processreport('s')
fillWeekReports(first_week, countdays)
processreport('w')
key = input("Wait")
