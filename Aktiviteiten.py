import os
import sys
import math
import unicodedata
from pathlib import Path
from datetime import datetime, date, timedelta
from ics import Calendar, Event
from reportlab.graphics import renderPDF
from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.colors import HexColor
from reportlab.lib.pagesizes import LETTER, A4, landscape, portrait
from reportlab.lib.units import inch, mm
from reportlab.lib.colors import blue, green, black, red, pink, gray, brown, purple, orange, yellow, white
from reportlab.pdfbase import pdfmetrics  
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import registerFontFamily
from svglib.svglib import svg2rlg, load_svg_file, SvgRenderer

startdate = datetime(1990,1,1)
datecal = datetime.now()
activityfont = "LiberationSerif"
version = "December 2025"
weekdaynames = ["Maandag","Dinsdag","Woensdag","Donderdag","Vrijdag","Zaterdag","Zondag"]
monthnames = ["Januari","Februari","Maart","April","Mei","Juni","Juli","Augustus", "September","Oktober","November","December"]
activity_summary_x = 5
activity_summary_y = 50
activity_summary_dy = 10

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
        
def find_all_occurrences(line, sub, f, t):
    index_of_occurrences = []
    current_index = f
    while True:
        current_index = line.find(sub, current_index)
        if current_index == -1 or current_index >= t:
            return index_of_occurrences
        else:
            index_of_occurrences.append(current_index)
            current_index += len(sub)
        
def scaleSVG(svgfile, scaling_factor):
    svg_root = load_svg_file(svgfile)
    svgRenderer = SvgRenderer(svgfile)
    drawing = svgRenderer.render(svg_root)
    scaling_x = scaling_factor
    scaling_y = scaling_factor
    drawing.width = drawing.minWidth() * scaling_x
    drawing.height = drawing.height * scaling_y
    drawing.scale(scaling_x, scaling_y)
    return drawing
        
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
    img = ""
    newcalendar = [[] for _ in range(300)]
    newcalendar[3] = "kerkdienst"
    newcalendar[12] = "optredenklassiek"
    newcalendar[14] = "bibliotheek"
    newcalendar[16] = "uiteten"
    newcalendar[30] = "wandelen"
    newcalendar[31] = "sportenspelmiddag"
    newcalendar[32] = "wenskaarten"
    newcalendar[33] = "bloemschikken"
    newcalendar[34] = "klassiekemuziek"
    newcalendar[88] = "sjoelen"
    img = newcalendar[int(imgcode)]
    return img

def processsdescription(text):
    imgcode = ""
    imgpos = ''
    dtimgeventpos = text.find("[i")
    if dtimgeventpos >= 0:
        imgcode = text[dtimgeventpos + 2:dtimgeventpos + 5]
        if imgcode[1] == ']':
            imgcode = "00" + imgcode[0:1]
        if imgcode[2] == ']':
            imgcode = "0" + imgcode[0:2]
    return imgcode
    
def breakoff(textarray, font, fontsize, limitlength):
    first = textarray[0]
    rest = []
    for j in range(1, len(textarray)):
        summarywidth = pdfmetrics.stringWidth(first + " " + textarray[j], activityfont, 10)
        if summarywidth < 100:
            first = first + " " + textarray[j]
        else:
            rest = textarray[j:len(textarray)]
            break
    return (first, rest)
    
def drawActivity(c, activity_x, activity_y, w, h, a, i):
    c.setFillColor(HexColor(whitelayover))
    p = c.beginPath()
    p.moveTo(activity_x, activity_y + 0.5 * a)
    p.arcTo(activity_x, activity_y, activity_x + a, activity_y + a, startAng = 180, extent = 90)  # arc left below
    p.lineTo(activity_x + w, activity_y)                                                           # horizontal line
    p.arcTo(activity_x + w, activity_y, activity_x + w + a, activity_y + a, startAng = 270, extent = 90)  # arc right below
    p.lineTo(activity_x + w + a, activity_y + h)                                                      # vertcal line
    p.arcTo(activity_x + w, activity_y + h, activity_x + w + a, activity_y + h + a, startAng = 0, extent = 90)     # arc right above
    p.lineTo(activity_x + 0.5 * a, activity_y + h + a)                                                   # horizontal line
    p.arcTo(activity_x, activity_y + h, activity_x + a, activity_y + h + a, startAng = 90, extent = 90)    # arc left above
    p.lineTo(activity_x, activity_y + 0.5 * a)                                                                # vertcal line
    c.drawPath(p, stroke = 0, fill = 1)
    drawing = scaleSVG("SVG/location.svg", 0.02)
    renderPDF.draw(drawing, c, activity_x + 5, activity_y + 5)
    daytimestr = str(monthevents[i].day) + " " + weekdaynames[monthevents[i].weekday] + " " + monthevents[i].starttime + "-" + monthevents[i].endtime
    c.setFillColor(HexColor(blacktext))
    c.setFont(activityfont, 12)
    c.drawString(activity_x + 5, activity_y + 70, daytimestr)
    c.setFont(activityfont, 10)
    inparts = monthevents[i].summary.split()
    (first, rest) = breakoff(inparts, activityfont, 10, 150)
    c.drawString(activity_x + activity_summary_x, activity_y + activity_summary_y, first)
    if len(rest) > 0:
        (first, rest) = breakoff(rest, activityfont, 10, 150)
        c.drawString(activity_x + activity_summary_x, activity_y + activity_summary_y - activity_summary_dy, first)
    imgcode = processsdescription(monthevents[i].description)
    activity_kind_x = 75
    activity_kind_y = 100
    activity_kind_r = 20
    if len(imgcode) > 0:
        img = lookupimage(imgcode)
        if len(img) > 0:
            c.radialGradient(activity_x + activity_kind_x, activity_y + activity_kind_y, activity_kind_r, (pinkredcircle, orangecircle), extend = False)
            c.circle(activity_x + activity_kind_x, activity_y + activity_kind_y, activity_kind_r, stroke = 0)
            drawing = scaleSVG("SVG/" + img + ".svg", 0.2)
            renderPDF.draw(drawing, c, activity_x + activity_kind_x - 12, activity_y + activity_kind_y - 12)
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

A4_width = A4[0]
A4_height = A4[1]

yellowbackground = "#ffde22"
lighteryellow = "#ffec60"
pinkredcircle = "#ff414e"
orangecircle = "#ff8928"
whitelayover = "#ffffff"
blacktext = "#000000"

activity_x = 100
activity_y = 200
activity_w = 200
activity_h = 200
activity_kind_w = 100
activity_kind_h = 100
activity_kind_r = 0.5 * activity_kind_w

c = Canvas("PDF/Aktiviteiten.pdf", pagesize=landscape(A4))
c.setFillColor(HexColor(yellowbackground))
c.rect(0, 0, A4_height, A4_width, fill = 1)
c.setFillColor(HexColor(lighteryellow))
c.rect(75, 95, 300, 200, stroke = 0, fill = 1)
activity_x = 60
activity_y = 440
col = 0
for i in range(0, 16):
    drawActivity(c, activity_x,  activity_y, 130, 70, 20, i)
    col += 1
    activity_x = activity_x + 180
    if col == 4:
        col = 0
        activity_x = 60
        activity_y = activity_y - 130
c.showPage()
c.setFillColor(HexColor(yellowbackground))
c.rect(0, 0, A4_height, A4_width, fill = 1)
c.setFillColor(HexColor(lighteryellow))
c.rect(75, 95, 300, 200, stroke = 0, fill = 1)
activity_x = 60
activity_y = 440
col = 0
for i in range(16, 32):
    drawActivity(c, activity_x,  activity_y, 130, 70, 20, i)
    col += 1
    activity_x = activity_x + 180
    if col == 4:
        col = 0
        activity_x = 60
        activity_y = activity_y - 130
c.showPage()
c.setFillColor(HexColor(yellowbackground))
c.rect(0, 0, A4_height, A4_width, fill = 1)
c.setFillColor(HexColor(lighteryellow))
c.rect(75, 95, 300, 200, stroke = 0, fill = 1)
activity_x = 60
activity_y = 440
col = 0
for i in range(32, len(monthevents)):
    drawActivity(c, activity_x,  activity_y, 130, 70, 20, i)
    col += 1
    activity_x = activity_x + 180
    if col == 4:
        col = 0
        activity_x = 60
        activity_y = activity_y - 130
c.showPage()
c.save()

key = input("Wait")
