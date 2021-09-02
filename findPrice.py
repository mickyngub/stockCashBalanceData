import openpyxl
import datetime
from openpyxl.utils.cell import get_column_letter


workbook = openpyxl.load_workbook(filename="Main.xlsx")
sheet = workbook.active
number_rows = sheet.max_row
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
            
            
            sheet["F"+str(i+1)] = float(tickerCPEndDateM5) if tickerCPEndDateM5 != "None" else "DNE"
            sheet["G"+str(i+1)] = float(tickerCPEndDateM4) if tickerCPEndDateM4 != "None" else "DNE"
            sheet["H"+str(i+1)] = float(tickerCPEndDateM3) if tickerCPEndDateM3 != "None" else "DNE"
            sheet["I"+str(i+1)] = float(tickerCPEndDateM2) if tickerCPEndDateM2 != "None" else "DNE"
            sheet["J"+str(i+1)] = float(tickerCPEndDateM1) if tickerCPEndDateM1 != "None" else "DNE"
            sheet["K"+str(i+1)] = float(tickerCPEndDate) if tickerCPEndDate != "None" else "DNE"     
            sheet["L"+str(i+1)] = float(tickerCPEndDateP1) if tickerCPEndDateP1 != "None" else "DNE"
            sheet["M"+str(i+1)] = float(tickerCPEndDateP2) if tickerCPEndDateP2 != "None" else "DNE"
            sheet["N"+str(i+1)] = float(tickerCPEndDateP3) if tickerCPEndDateP3 != "None" else "DNE"
            sheet["O"+str(i+1)] = float(tickerCPEndDateP4) if tickerCPEndDateP4 != "None" else "DNE"
            sheet["P"+str(i+1)] = float(tickerCPEndDateP5) if tickerCPEndDateP5 != "None" else "DNE"

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
            workbook.save("Main.xlsx")

            break
        k += 1
    print("-------------------------------")

workbook.save("Main.xlsx")