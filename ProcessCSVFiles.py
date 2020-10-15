import os
import pandas as pd
import xlwt
from xlwt import Workbook


# This function is responsible to process all csv files in specific directory and saving the processed results in an excel file.
def processCSVFiles():
    directory = os.path.join("c:\\","Users\\faqeerrehman\\MSU\\Sajid")
    resultFile = "C:\\Users\\faqeerrehman\\MSU\\Softwares\\Results.xls"
    # Workbook is created
    wb = Workbook()
    # add_sheet is used to create sheet.
    sheet1 = wb.add_sheet('Sheet 1')
    #Add column names
    sheet1.write(0, 0, "DateTime")
    sheet1.write(0, 1, "TravelTime")

    rowToBeWritten = 1;
    for root,dirs,files in os.walk(directory):
        for file in files:
           if file.endswith(".csv"):
               filename = os.path.join(directory,file)
               data = pd.read_csv(filename)
               splittedData = data["Timestamp;TravelTimeMs;LocalTimeStamp"].str.split(";", expand = True)
               averageTravelTimeMilliseconds = sum(splittedData[1].astype(int)) / len(splittedData[1].astype(int))
               averageTime = returnFormattedTime(averageTravelTimeMilliseconds)
               PerDayRowValue = splittedData[2].iloc[0]
               sheet1.write(rowToBeWritten, 0, PerDayRowValue)
               sheet1.write(rowToBeWritten, 1, averageTime)
               rowToBeWritten = rowToBeWritten + 1
        wb.save(resultFile)

#This function is reponsible to get travel time (in milliseconds) as input and return it in the form of hours:minuts:seconds
def returnFormattedTime(averageTravelTimeMilliseconds = None):
    averageTravelTimeSeconds = (averageTravelTimeMilliseconds / 1000) % 60
    averageTravelTimeMinutes = (averageTravelTimeMilliseconds / (1000 * 60)) % 60
    averageTravelTimeHours = (averageTravelTimeMilliseconds / (1000 * 60 * 60)) % 24
    formattedAverageTravelTime = "%d:%d:%d" % (averageTravelTimeHours, averageTravelTimeMinutes, averageTravelTimeSeconds)
    return formattedAverageTravelTime

#Main function (entry point of program)
if __name__ == '__main__':
    processCSVFiles()