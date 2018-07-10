from datetime import timedelta, date, time
import datetime

def convert24(str1):
        # Checking if last two elements of time
        # is AM and first two elements are 12
        if str1[-2:] == "AM" and str1[:2] == "12":
            hour = "0"
            minute = str1[3:-3]
            return int(hour), int(minute)   
        # remove the AM    
        elif str1[-2:] == "AM":
            if str1[0] == "0":
                hour = str1[1]
            else:
                hour = str1[:2]
            if str1[3] == "0":
                minute = str1[4]
            else:
                minute = str1[3:-3]
            
            #minute =  str1[3:-3]
            return int(hour), int(minute)
         
        # Checking if last two elements of time
        # is PM and first two elements are 12   
        elif str1[-2:] == "PM" and str1[:2] == "12":
            if str1[0] == "0":
                hour = str1[1]
            else:
                hour = str1[:2]
            if str1[3] == "0":
                minute = str1[4]
            else:
                minute = str1[3:-3]
            return int(hour), int(minute)
             
        else:
             
            # add 12 to hours and remove PM
            hour = int(str1[:2]) + 12
            
            if str1[3] == "0":
                minute = str1[4]
            else:
                minute = str1[3:-3]

            return int(hour), int(minute)
 
def flaggedWords(ws, my_list, results_list):
    '''Finds keywords in row of data, throws in list'''
    for row in ws.iter_rows(row_offset=1):
        d = row[0] # The note
        e = row[1] # The name of the individual
        f = row[2] # Contact date
        g = row[3] # Program
        h = row[4] # Start time
        i = row[5] # end time
        j = row[6] # duration
        k = row[7] # Note writer
        foundWords = []
        if d.value:
            for w in sorted(my_list):
                if w in str(d.value).lower():
                    foundWords.append(w)
            if len(foundWords) > 0:
                note = ''
                for l in foundWords:
                    left,sep,right = d.value.lower().partition(l)
                    note = note + "..." + left[-70:] + sep.upper() + right[:70] + "..." + ';'
                forCSV = ','.join(foundWords).upper()
                this_list = [forCSV, e.value, h.value, i.value, f.value, note, g.value, j.value, k.value]
                results_list.append(this_list)
    return            
                #self.csvWritee(forCSV, e, h, i, f, note, g, j, k)
def flaggedWordsInverse(ws, my_list, results_list):
    '''Finds keywords in row of data, throws in list'''
    for row in ws.iter_rows(row_offset=1):
        d = row[0] # The note
        e = row[1] # The name of the individual
        f = row[2] # Contact date
        g = row[3] # Program
        h = row[4] # Start time
        i = row[5] # end time
        j = row[6] # duration
        k = row[7] # Note writer
        foundWords = []
        if d.value:
            for w in sorted(my_list):
                if w not in str(d.value).lower():
                    foundWords.append(w)
            if len(foundWords) > 0:
                note = ''
                for l in foundWords:
                    left,sep,right = d.value.lower().partition(l)
                    note = note + left[-70:] + sep.upper() + right[:70] + ';'
                forCSV = ','.join(foundWords).upper()
                this_list = ['Missing ' + forCSV, e.value, h.value, i.value, f.value, note, g.value, j.value, k.value]
                results_list.append(this_list)
                
                #self.csvWritee(forCSV, e, h, i, f, note, g, j, k)
    return
def oddDuration(ws, greaterthan, lessthan, results_list):
    for row in ws.iter_rows(row_offset=1):
        d = row[0] # The note
        e = row[1] # The name of the individual
        f = row[2] # Contact date
        g = row[3] # Program
        h = row[4] # Start time
        i = row[5] # end time
        j = row[6] # duration
        k = row[7] # Note writer
        if d.value:
            if j.value:
                if j.value <= lessthan or j.value >= greaterthan:
                    note = d.value.split('.')
                    d = '.'.join(note[1:3]).lstrip() + ' [. . .] ' + d.value[-100:]
                    this_list = ['Duration', e.value, h.value, i.value, f.value, d, g.value, j.value, k.value]
                    results_list.append(this_list)

                    #self.csvWritee('Duration', e, h, i, f, d, g, j, k)
            else:
                this_list = ['No Duration', e.value, h.value, i.value, f.value, d.value, g.value, j.value, k.value]
                results_list.append(this_list)
                #self.csvWritee('NO DURATION', e, h, i, f, d.value, g, j, k)
    return
def shortNote(ws, notelength, results_list):

    for row in ws.iter_rows(row_offset=1):
        d = row[0] # The note
        e = row[1] # The name of the individual
        f = row[2] # Contact date
        g = row[3] # Program
        h = row[4] # Start time
        i = row[5] # end time
        j = row[6] # duration
        k = row[7] # Note writer
        if d.value:
            if len(d.value) < notelength:
                this_list = ['SHORT NOTE (<' + str(notelength) + ')', e.value, h.value, i.value, f.value, d.value, g.value, j.value, k.value]
                results_list.append(this_list)
                #self.csvWritee('SHORT NOTE (< ' + str(notelength) + ')', e, h, i, f, d.value, g, j, k)
    return
def oddTimes(ws, startTimeAfter, startTimeBefore, results_list):

    after = convert24(startTimeAfter)
    afterHour = after[0]
    afterMin = after[1]
    before = convert24(startTimeBefore)
    beforeHour = before[0]
    beforeMin = before[1]
    for row in ws.iter_rows(row_offset=1):
        d = row[0] # The note
        e = row[1] # The name of the individual
        f = row[2] # Contact date
        g = row[3] # Program
        h = row[4] # Start time
        i = row[5] # end time
        j = row[6] # duration
        k = row[7] # Note writer
        if h.value:
            note = d.value.split('.')
            d = '.'.join(note[1:3]).lstrip() + ' [. . .] ' + d.value[-100:]
            try:
                if h.value > time(afterHour, afterMin):
                    #self.csvWritee("START TIME AFTER " + startTimeAfter, e, h, i, f, d, g, j, k)
                    this_list = ['START TIME AFTER ' + startTimeAfter, e.value, h.value, i.value, f.value, d, g.value, j.value, k.value]
                    results_list.append(this_list)
                elif h.value < time(beforeHour, beforeMin):
                    this_list = ['START TIME BEFORE ' + startTimeBefore, e.value, h.value, i.value, f.value, d, g.value, j.value, k.value]
                    results_list.append(this_list)
                    #self.csvWritee("START TIME Before " + startTimeBefore, e, h, i, f, d, g, j, k)
            except (TypeError):
                this_list = ['12AM/Error', e.value, h.value, i.value, f.value, d, g.value, j.value, k.value]
                results_list.append(this_list)
                #self.csvWritee("12AM/Error", e, h, i, f, d, g, j, k)
    return
def underUnits(ws, underUnits, results_list):
    units = int(underUnits) * 15
    from collections import defaultdict
    names = defaultdict(int)
    for row in ws.iter_rows(row_offset=1):
        d = row[0] # The note
        e = row[1] # The name of the individual
        f = row[2] # Contact date
        g = row[3] # Program
        h = row[4] # Start time
        i = row[5] # end time
        j = row[6] # duration
        k = row[7] # Note writer

        if j.value is None:
            pass
        else:
            names[e.value] += j.value
    for k, v in names.items():
        if names[k] < units:
            this_list = ['UNDER UNITS (' +str(underUnits) + ')', k, str(int(v)/15), '', '','', '', '', '', '']
            results_list.append(this_list)
'''
def under7forResidential(self):
        example_dictionary = defaultdict(list)
        for row in ws:
            a = row[0]
            b = row[1]
            if a.value:
                example_dictionary[a.value].append(b.value)

        namesClean = {}

        for names, values in example_dictionary.items():
            namesClean[names] = set(values)

        for names, values in namesClean.items():
            if len(values) < 2:
                print(names + ' has less than two dates')

'''