import datetime as dt
import dateutil.parser as dparser
import xlrd, os

totalUnanswered = 0 
totalResponseTime = 0.0
totalResponses = 0
totalCSLCompletes = 0
totalCSLCases = 0

def loopThroughFiles():
    global totalUnanswered
    global totalResponseTime
    global totalResponses
    
    for file in os.listdir('files'):
        if not (file.endswith(".xls") or file.endswith("xlsx")):
            continue
        
        # Open the workbook
        workbook = xlrd.open_workbook("files/" + file)
        
        if('mydtxt' in file.lower()):
            calculateDeezFeedbackTime(workbook)
            continue
        elif('csl' in file.lower()):
            calculateCSLCompleteness(workbook)
            continue
        
    
    totalAverageResponseTime = totalResponseTime / totalResponses
    cslCompleteRate = float(totalCSLCompletes) / float(totalCSLCases) * 100
    
    print("Average Response Time: %.3f hours" % totalAverageResponseTime)
    print("Total unanswered: %d" % totalUnanswered)
    print("CSL Complete Rate: %.2f%%" % cslCompleteRate)

    
def calculateCSLCompleteness(workbook):
    global totalCSLCompletes
    global totalCSLCases
    
    # Loop over all of the sheets in the workbook
    for i in range (0, len(workbook.sheet_names())):
        # Open the current sheet
        worksheet = workbook.sheet_by_index(i)

        # The starting row since the title rows mean nothing
        row = 2
        
        while(row < worksheet.nrows):
            if(worksheet.cell(row, 0).value == ""):
                row += 1
                continue
            
            if(worksheet.cell(row, 8).value.lower() == 'y'):
                totalCSLCompletes += 1
            totalCSLCases += 1
            row += 1
        
    
def calculateDeezFeedbackTime(workbook):
    global totalUnanswered
    global totalResponseTime
    global totalResponses
    

    # Open the first sheet
    worksheet = workbook.sheet_by_index(0)

    # The starting row since the title rows mean nothing
    row = 2

    while(row < worksheet.nrows):
        receivedTime = worksheet.cell(row, 0).value

        if(receivedTime == ""):
            row += 1
            continue

        responseSent = worksheet.cell(row, 2).value

        if(responseSent == ""):
            totalUnanswered += 1
            row += 1
            continue

        responseTime = calculateDifference(receivedTime, responseSent)

        if(responseTime != -1):
            totalResponseTime += calculateDifference(receivedTime, responseSent)
            totalResponses += 1
        row += 1
    
    
def calculateDifference(time1, time2):
    try:
        date=dt.datetime.strptime(time1,'%m/%d/%Y %I:%M:%S %p')
        date2=dt.datetime.strptime(time2,'%m/%d/%Y %H:%M')
    except (TypeError, ValueError):
        return -1
    
    difference = date2 - date
    difInHours =  (difference.total_seconds() / 60) / 60
    #print(date)
    #print(date2)
    #print("%.3f\n" % difInHours)
    return difInHours

loopThroughFiles()