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
version = "November 2025"
weekdaynames = ["Maandag","Dinsdag","Woensdag","Donderdag","Vrijdag","Zaterdag","Zondag"]
monthnames = ["Januari","Februari","Maart","April","Mei","Juni","Juli","Augustus", "September","Oktober","November","December"]

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

pdfmetrics.registerFont(TTFont('Ubuntu', 'Ubuntu-Regular.ttf'))
pdfmetrics.registerFont(TTFont('UbuntuBold', 'Ubuntu-Bold.ttf'))
pdfmetrics.registerFont(TTFont('UbuntuItalic', 'Ubuntu-Italic.ttf'))
pdfmetrics.registerFont(TTFont('UbuntuBoldItalic', 'Ubuntu-BoldItalic.ttf'))
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

c = Canvas("PDF/Aktiviteiten.pdf")
c.showPage()
c.save()

key = input("Wait")
