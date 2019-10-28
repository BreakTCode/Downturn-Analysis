import pandas as pd
from datetime import date as dt
import xlsxwriter as xlw

#Functions
def Add_to_Worksheet(worksheet, sheet, i, count):
    myVals = [
                    str(pd.to_datetime(sheet['Date'][i])), 
                    str(sheet['Close'][i])   
                ]
    worksheet.write(('A' + str(count)), i)            
    worksheet.write_row(('B' + str(count)), myVals)

    return count + 1

sheets = pd.read_excel('^GSPC 1974.xlsx', header=0)
writer = xlw.Workbook('Downdraw Instances 1974 ' + str(dt.today()) + '.xlsx')
worksheet1 = writer.add_worksheet('Downturn Instances 5%')
worksheet2 = writer.add_worksheet('Downturn Instances 10%')
worksheet3 = writer.add_worksheet('Downturn Instances 20%')

myheaders = ['Row', 'Date', 'Close']
worksheet1.write_row('A1', myheaders)
worksheet2.write_row('A1', myheaders)
worksheet3.write_row('A1', myheaders)

weeks = 0
Total5Downs = 0
peak = 0
maxdown = 0
DayRecovered = 0
Average = 0
recovered = True
DTCount1 = 2
DTCount2 = 2
DTCount3 = 2
i = 1
downspot = 0

while i < sheets.shape[0]:
    if ((sheets['Date'][i] - sheets['Date'][i - 1]).days) >= 3:
        weeks += 1
    i += 1
i = 0
while i < sheets.shape[0]:
    #Check if this is the first row.
    if peak == 0:
        peak = sheets['Close'][0]
        DayRecovered = sheets['Date'][0].day
    #Check if our current value is greater than our peak.
    if peak <= sheets['Close'][i]:
        peak = sheets['Close'][i]
        downspot = peak
    #Check if our value is less than 5% of peak
   
    elif sheets['Close'][i] <= downspot * .95:
        maxdown = sheets['Close'][i]
        if sheets['Close'][i] <= peak *0.95:
            downspot = sheets['Close'][i]
            Total5Downs +=1
            recovered = False
        if sheets['Close'][i] <= peak *0.90:
            downspot = sheets['Close'][i]
            Total5Downs +=1
            recovered = False
        if sheets['Close'][i] <= peak *0.85:
            downspot = sheets['Close'][i]
            Total5Downs +=1
            recovered = False
        if sheets['Close'][i] <= peak *0.80:
            downspot = sheets['Close'][i]
            Total5Downs +=1
            recovered = False
        if sheets['Close'][i] <= peak *0.75:
            downspot = sheets['Close'][i]
            Total5Downs +=1
        if sheets['Close'][i] <= peak *0.70:
            downspot = sheets['Close'][i]
            Total5Downs +=1
        if sheets['Close'][i] <= peak *0.65:
            downspot = sheets['Close'][i]
            Total5Downs +=1
        if sheets['Close'][i] <= peak *0.60:
            downspot = sheets['Close'][i]
            Total5Downs +=1
        if sheets['Close'][i] <= peak *0.55:
            downspot = sheets['Close'][i]
            Total5Downs +=1
        if sheets['Close'][i] <= peak *0.50:
            downspot = sheets['Close'][i]
            Total5Downs +=1
        if sheets['Close'][i] <= peak *0.45:
            downspot = sheets['Close'][i]
            Total5Downs +=1
        
        recovered = False
        Average = Average + abs(sheets['Date'][i].day - DayRecovered)
        DTCount1 = Add_to_Worksheet(worksheet1, sheets, i, DTCount1)
    

    #Test if current value recovered half of the downturn
    elif ((peak - maxdown) / 2) <= sheets['Close'][i] - maxdown and recovered == False:
       
        recovered = True
        peak = sheets['Close'][i]
        downspot = peak
        DayRecovered = sheets['Date'][i].day

    i += 1

writer.close()
print(weeks)
print(Total5Downs)
print(weeks/Total5Downs)

