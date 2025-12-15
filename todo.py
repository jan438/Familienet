def breakoff(textarray, fontsize, limitlength):
    first = textarray[0]
    firstwidth = pdfmetrics.stringWidth(first,  activityfont, fontsize)  
    margin = limitlength - firstwidth
    rest = []
    for j in range(1, len(textarray)):
        firstwidth = pdfmetrics.stringWidth(first + " " + textarray[j], activityfont, fontsize)
        if firstwidth < limitlength:
        	margin = limitlength - firstwidth
            first = first + " " + textarray[j]
        else:
            rest = textarray[j:len(textarray)]
            break
    return (first, rest, margin)