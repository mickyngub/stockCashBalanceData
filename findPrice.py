import openpyxl
import datetime
from openpyxl.utils.cell import get_column_letter


workbook = openpyxl.load_workbook(filename="Main.xlsx")
sheet = workbook.active
number_rows = 3
number_columns = sheet.max_column



for i in range(1,number_rows):
    print("-------------------------------")
    tickerName = str(sheet[("A"+str(i+1))].value)
    endDateValue = str(sheet[("D"+str(i+1))].value)[:10]
    endDate = datetime.datetime.strptime(endDateValue, '%Y-%m-%d')
    # endDateM5 = endDate + datetime.timedelta(days=-5)

    strEndDate = endDate.strftime('%d/%m/%Y')
    # strEndDateM5 = endDateM5.strftime('%d/%m/%Y')
   
    print(tickerName + " endDate " + strEndDate)
    
    tickerFileOpen = "cleanDate" + tickerName + ".xlsx"
    print(".....Opening '" + tickerFileOpen + "'....")

    tickerWorkBook = openpyxl.load_workbook(filename=tickerFileOpen)
    tickerSheet = tickerWorkBook.active
    tickerNumber_rows = tickerSheet.max_row
    # tickerNumber_columns = tickerSheet.max_column
    k = 0
    while k < tickerNumber_rows:
        tickerDate = str(tickerSheet["A"+str(k+1)].value)
        tickerCell = "A" + str(k+1)
        # print('this is tickerdate', tickerDate, ' strEndDate', strEndDate)
        if tickerDate == strEndDate:
            # print('Found EndDate at Cell Number'+tickerCell)

            tickerDateM5 = str(tickerSheet["A"+str(k-4)].value)
            tickerDateM4 = str(tickerSheet["A"+str(k-3)].value)
            tickerDateM3 = str(tickerSheet["A"+str(k-2)].value)
            tickerDateM2 = str(tickerSheet["A"+str(k-1)].value)
            tickerDateM1 = str(tickerSheet["A"+str(k)].value)
            tickerDate = str(tickerSheet["A"+str(k+1)].value)
            tickerDateP1 = str(tickerSheet["A"+str(k+2)].value)
            tickerDateP2 = str(tickerSheet["A"+str(k+3)].value)
            tickerDateP3 = str(tickerSheet["A"+str(k+4)].value)
            tickerDateP4 = str(tickerSheet["A"+str(k+5)].value)
            tickerDateP5 = str(tickerSheet["A"+str(k+6)].value)

            tickerCPEndDateM5 = str(tickerSheet["E"+str(k-4)].value)
            tickerCPEndDateM4 = str(tickerSheet["E"+str(k-3)].value)
            tickerCPEndDateM3 = str(tickerSheet["E"+str(k-2)].value)
            tickerCPEndDateM2 = str(tickerSheet["E"+str(k-1)].value)
            tickerCPEndDateM1 = str(tickerSheet["E"+str(k)].value)
            tickerCPEndDate = str(tickerSheet["E"+str(k+1)].value)
            tickerCPEndDateP1 = str(tickerSheet["E"+str(k+2)].value)
            tickerCPEndDateP2 = str(tickerSheet["E"+str(k+3)].value)
            tickerCPEndDateP3 = str(tickerSheet["E"+str(k+4)].value)
            tickerCPEndDateP4 = str(tickerSheet["E"+str(k+5)].value)
            tickerCPEndDateP5 = str(tickerSheet["E"+str(k+6)].value)
            
            sheet["F"+str(i+1)] = tickerCPEndDateM5
            sheet["G"+str(i+1)] = tickerCPEndDateM4
            sheet["H"+str(i+1)] = tickerCPEndDateM3
            sheet["I"+str(i+1)] = tickerCPEndDateM2
            sheet["J"+str(i+1)] = tickerCPEndDateM1
            sheet["K"+str(i+1)] = tickerCPEndDate
            sheet["L"+str(i+1)] = tickerCPEndDateP1
            sheet["M"+str(i+1)] = tickerCPEndDateP2
            sheet["N"+str(i+1)] = tickerCPEndDateP3
            sheet["O"+str(i+1)] = tickerCPEndDateP4
            sheet["P"+str(i+1)] = tickerCPEndDateP5

            print(tickerName + " " + tickerDateM5 + " " + tickerCPEndDateM5)
            print(tickerName + " " + tickerDateM4 + " " + tickerCPEndDateM4)
            print(tickerName + " " + tickerDateM3 + " " + tickerCPEndDateM3)
            print(tickerName + " " + tickerDateM2 + " " + tickerCPEndDateM2)
            print(tickerName + " " + tickerDateM1 + " " + tickerCPEndDateM1)
            print("END CashBalance " + tickerName + " " + tickerDate + " " + tickerCPEndDate)
            print(tickerName + " " + tickerDateP1 + " " + tickerCPEndDateP1)
            print(tickerName + " " + tickerDateP2 + " " + tickerCPEndDateP2)
            print(tickerName + " " + tickerDateP3 + " " + tickerCPEndDateP3)
            print(tickerName + " " + tickerDateP4 + " " + tickerCPEndDateP4)

            break
        k += 1
    print("-------------------------------")

workbook.save("Main.xlsx")