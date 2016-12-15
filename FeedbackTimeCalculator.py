import datetime as dt
import dateutil.parser as dparser
import xlrd, os

totalUnanswered = 0 
totalResponseTime = 0.0
totalResponses = 0

def loopThroughFiles():
    global totalUnanswered
    global totalResponseTime
    global totalResponses
    
    for file in os.listdir('files'):
        if not (file.endswith(".xls") or file.endswith("xlsx")):
            continue
        
        # Open the workbook
        workbook = xlrd.open_workbook("files/" + file)

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
    
    totalAverageResponseTime = totalResponseTime / totalResponses
    
    print("Average Response Time: %.3f hours" % totalAverageResponseTime)

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