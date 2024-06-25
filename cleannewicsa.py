import pytz
import os
import sys
import shutil
from pathlib import Path

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
            
def process_alarm(line, pos, endpos):
    print("Alarm", pos, endpos, line[pos:endpos+12])
    processed = line[:pos] + line[endpos + 12:]
    return processed

def process_alarms(line, pos):
    processed = line
    print("count alarms", len(pos))
    for i in range(len(pos), 0, -1):
        endpos = processed.find(ENDVALARM, pos[i-1])
        print(i, "Process alarm", processed, "i", i, "start", pos[i-1], "end", endpos)
        processed = process_alarm(processed, pos[i-1], endpos)
    return processed
            
def process_organizer(line, pos):
    processed = line[:pos] + line[pos+36:]
    return processed
    
def process_organizers(line, pos):
    processed = line
    for i in range(len(pos), 0, -1):
        processed = process_organizer(processed, pos[i-1])
    return processed
           
def process_linebreak(line, pos):
    processed = line[:pos] + line[pos+3:]
    return processed
    
def process_linebreaks(line, pos):
    processed = line
    for i in range(len(pos), 0, -1):
        processed = process_linebreak(processed, pos[i-1])
    return processed
    
def process_backslash(line, pos):
    processed = line[:pos] + line[pos+1:]
    return processed

if sys.platform[0] == 'l':
    inppath = '/home/jan/git/Familienet/ICS'
    outpath = '/home/jan/git/Familienet/Calendar'
if sys.platform[0] == 'w':
    inppath = "C:/Users/janbo/OneDrive/Documents/GitHub/Familienet/ICS"
    outpath = "C:/Users/janbo/OneDrive/Documents/GitHub/Familienet/Calendar"
os.chdir(inppath)
#src = "NewCalendar.ics"
src = "FamilienetAlarm.ics"
#src = "KOrganizer.ics"
bup = "Familienet.bup"
cleaned = "Familienet.ics"
shutil.copy(src, bup)
inpfile = open(bup, 'rb')
outfile = open(cleaned, 'wb')
line = inpfile.read()
BEGINVEVENT = "BEGIN:VEVENT".encode()
BEGINVALARM = "BEGIN:VALARM".encode()
ENDVALARM = "END:VALARM".encode()
SUMMARY = "SUMMARY".encode()
DESCRIPTION = "DESCRIPTION".encode()
LOCATION = "LOCATION".encode()
DTSTART = "DTSTART".encode()
DTEND = "DTEND".encode()
crlf = '\r\n'.encode()
linebreak = '\r\n '.encode()
backslash = "\\".encode()
ORGANIZER = "ORGANIZER:mailto:local@newcalendar".encode()
print("============", line, len(line))
alarms = find_all_occurrences(line, BEGINVALARM, 0, len(line))
line = process_alarms(line, alarms)
organizers = find_all_occurrences(line, ORGANIZER, 0, len(line))
line = process_organizers(line, organizers)
neweventpos = line.find(BEGINVEVENT)
dtstartpos = line.find(DTSTART)
dtendpos = line.find(DTEND)
summarypos = line.find(SUMMARY)
descriptionpos = line.find(DESCRIPTION)
locationpos = line.find(LOCATION)
summaries = find_all_occurrences(line, SUMMARY, 0, len(line))
descriptions = []
for i in range(len(summaries)):
    descriptionpos = line.find(DESCRIPTION, summaries[i])
    descriptions.append(descriptionpos)
for i in range(len(summaries)):
    linebreaks = find_all_occurrences(line, linebreak, summaries[i], descriptions[i])
    if len(linebreaks) > 0:
        line = process_linebreaks(line, linebreaks)
descriptions = find_all_occurrences(line, DESCRIPTION, 0, len(line))
locations = []
for i in range(len(descriptions)):
    locationpos = line.find(LOCATION, descriptions[i])
    locations.append(locationpos)
for i in range(len(descriptions)):
    linebreaks = find_all_occurrences(line, linebreak, descriptions[i], locations[i])
    if len(linebreaks) > 0:
        line = process_linebreaks(line, linebreaks)
    backslashpos = line.find(backslash)
    while backslashpos > 0:
        line = process_backslash(line, backslashpos)
        backslashpos = line.find(backslash)
print("============", line, len(line))
outfile.write(line)
inpfile.close()
outfile.close()
