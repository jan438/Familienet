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
    alarmtext = line[pos:] + line[:endpos]
    decodedtext = alarmtext.decode("utf-8")
    alarmtime = decodedtext[25:27]
    if alarmtime == "60":
        processed = line[:pos] + "A[60]".encode() + crlf + line[endpos+12:]
    elif alarmtime == "15":
        processed = line[:pos] + "A[15]".encode() + crlf + line[endpos+12:]
    elif alarmtime == "14":
        processed = line[:pos] + "M[14]".encode() + crlf + line[endpos+12:]
    elif alarmtime == "10":
        processed = line[:pos] + "M[10]".encode() + crlf + line[endpos+12:]
    else:
        processed = line[:pos] + "A[60]".encode() + crlf + line[endpos+12:]
    return processed

def process_alarms(line, pos):
    processed = line
    for i in range(len(pos), 0, -1):
        endpos = processed.find(ENDVALARM, pos[i-1])
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
    
def process_2symbol(line, pos, wsymbol):
    processed = line[:pos] + wsymbol + line[pos+2:]
    return processed

def process_2symbols(line, pos, wsymbol):
    processed = line
    for i in range(len(pos), 0, -1):
        processed = process_2symbol(processed, pos[i-1], wsymbol)
    return processed

def process_symbol(line, pos, wsymbol):
    processed = line[:pos] + wsymbol + line[pos+3:]
    return processed

def process_symbols(line, pos, wsymbol):
    processed = line
    for i in range(len(pos), 0, -1):
        processed = process_symbol(processed, pos[i-1], wsymbol)
    return processed

def process_flag(line, pos):
    eb = line[pos+3]
    lb = line[pos+7]
    germancode = bytearray.fromhex("A9AA")
    francecode = bytearray.fromhex("ABB7")
    hollandcode = bytearray.fromhex("B3B1")
    spaincode = bytearray.fromhex("AAB8")
    austriacode = bytearray.fromhex("A6B9")
    polandcode = bytearray.fromhex("B5B1")
    if (eb == germancode[0] and lb == germancode[1]):
        landcode = flagprefix + "DEU".encode()
    elif (eb == francecode[0] and lb == francecode[1]):
        landcode = flagprefix + "FRA".encode()
    elif (eb == hollandcode[0] and lb == hollandcode[1]):
        landcode = flagprefix + "HOL".encode()
    elif (eb == spaincode[0] and lb == spaincode[1]):
        landcode = flagprefix + "ESP".encode()
    elif (eb == austriacode[0] and lb == austriacode[1]):
        landcode = flagprefix + "AUT".encode()
    elif (eb == polandcode[0] and lb == polandcode[1]):
        landcode = flagprefix + "POL".encode()
    else:
        landcode = "".encode()
    processed = line[:pos] + landcode + line[pos+8:]
    return processed

def process_flags(line, pos):
    processed = line
    for i in range(len(pos), 0, -2):
        processed = process_flag(processed, pos[i-2])
    return processed
    
def process_emoji(line, pos):
    smbytes = line[pos:pos+4]
    utf8code = smbytes.decode('utf-8')
    unicode = utf8code.encode('unicode_escape')
    unistr = unicode. decode("ascii")
    l = len(unistr)
    emojicode = unistr[l-3:].encode("utf-8")
    processed = line[:pos] + emojiprefix + emojicode + line[pos+4:]
    return processed

def process_emojis(line, pos):
    processed = line
    for i in range(len(pos), 0, -1):
        processed = process_emoji(processed, pos[i-1])
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
    path = '/home/jan/git/Familienet/Calendar'
if sys.platform[0] == 'w':
    path = "C:/Users/janbo/OneDrive/Documents/GitHub/Familienet/Calendar"
os.chdir(path)
src = "NewCalendar.ics"
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
leuro = '\u20ac'.encode('utf-8')
weuro = bytearray.fromhex("80")
lbullet = '\u2022'.encode('utf-8')
wbullet = bytearray.fromhex("95")
lslquote = '\u2018'.encode('utf-8')
wslquote = bytearray.fromhex("91")
lsrquote = '\u2019'.encode('utf-8')
wsrquote = bytearray.fromhex("92")
ldlquote = '\u201c'.encode('utf-8')
wdlquote = bytearray.fromhex("93")
ldrquote = '\u201d'.encode('utf-8')
wdrquote = bytearray.fromhex("94")
# 0xC3 0xA9
leaigu = u"\xe9".encode()
weaigu = bytearray.fromhex("E9")
# 0xC3 0xA8
legrave = u"\xe8".encode()
wegrave = bytearray.fromhex("E8")
# 0xC3 0xAB
letrema = u"\xeb".encode()
wetrema = bytearray.fromhex("EB")
# 0xC3 0xAA
lecirconflex = u"\xea".encode()
wecirconflex = bytearray.fromhex("EA")
# 0xC2 0xB0
lgrad = u"\xb0".encode()
wgrad = bytearray.fromhex("B0")
# 0xC2 0xBD
lhalf = u"\xbd".encode()
whalf = bytearray.fromhex("BD")
# 0xC2 0xBC
lonequart = u"\xbc".encode()
wonequart = bytearray.fromhex("BC")
# 0xC2 0xBE
lthreequart = u"\xbe".encode()
wthreequart = bytearray.fromhex("BE")
flagcode = bytearray.fromhex("F09F87")
flagprefix = "[f".encode("utf-8")
emojicode = bytearray.fromhex("F09F")
emojiprefix = "[e".encode("utf-8")
alarms = find_all_occurrences(line, BEGINVALARM, 0, len(line))
flagcode = bytearray.fromhex("F09F87")
flagprefix = "[f".encode("utf-8")
emojicode = bytearray.fromhex("F09F")
emojiprefix = "[e".encode("utf-8")
line = process_alarms(line, alarms)
organizers = find_all_occurrences(line, ORGANIZER, 0, len(line))
line = process_organizers(line, organizers)
if sys.platform[0] == 'w':
    euros = find_all_occurrences(line, leuro, 0, len(line))
    line = process_symbols(line, euros, weuro)
    bullets = find_all_occurrences(line, lbullet, 0, len(line))
    line = process_symbols(line, bullets, wbullet)
    egraves = find_all_occurrences(line, legrave, 0, len(line))
    line = process_2symbols(line, egraves, wegrave)
    eaigus = find_all_occurrences(line, leaigu, 0, len(line))
    line = process_2symbols(line, eaigus, weaigu)
    ecirconflexes = find_all_occurrences(line, lecirconflex, 0, len(line))
    line = process_2symbols(line, ecirconflexes, wecirconflex)
    etremas = find_all_occurrences(line, letrema, 0, len(line))
    line = process_2symbols(line, etremas, wetrema)
    slquotes = find_all_occurrences(line, lslquote, 0, len(line))
    line = process_symbols(line, slquotes, wslquote)
    srquotes = find_all_occurrences(line, lsrquote, 0, len(line))
    line = process_symbols(line, srquotes, wsrquote)
    dlquotes = find_all_occurrences(line, ldlquote, 0, len(line))
    line = process_symbols(line, dlquotes, wdlquote)
    drquotes = find_all_occurrences(line, ldrquote, 0, len(line))
    line = process_symbols(line, drquotes, wdrquote)
    grads = find_all_occurrences(line, lgrad, 0, len(line))
    line = process_2symbols(line, grads, wgrad)
    halfs = find_all_occurrences(line, lhalf, 0, len(line))
    line = process_2symbols(line, halfs, whalf)
    onequarts = find_all_occurrences(line, lonequart, 0, len(line))
    line = process_2symbols(line, onequarts, wonequart)
    threequarts = find_all_occurrences(line, lthreequart, 0, len(line))
    line = process_2symbols(line, threequarts, wthreequart)
    flagcodes = find_all_occurrences(line, flagcode, 0, len(line))
    line = process_flags(line, flagcodes)
    emojicodes = find_all_occurrences(line, emojicode, 0, len(line))
    line = process_emojis(line, emojicodes)
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
outfile.write(line)
inpfile.close()
outfile.close()
key = input("Wait")
