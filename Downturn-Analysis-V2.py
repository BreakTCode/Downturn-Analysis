import pandas as pd
from datetime import date as dt
import xlsxwriter as xlw

#Functions
def Add_to_Worksheet(worksheet, sheet, i, count, dropper):
    myVals = [
                    str(pd.to_datetime(sheet['Date'][i])), 
                    str(sheet['Close'][i]),
                    dropper   
                ]
    worksheet.write(('A' + str(count)), i)            
    worksheet.write_row(('B' + str(count)), myVals)

    return count + 1

sheets = pd.read_excel('^DJI.xlsx', header=0)
writer = xlw.Workbook('Downdraw Instances Dow ' + str(dt.today()) + '.xlsx')

worksheet1 = writer.add_worksheet('Downturn Instances 5%')
worksheet2 = writer.add_worksheet('Downturn Instances 10%')
worksheet3 = writer.add_worksheet('Downturn Instances 20%')

myheaders = ['Row', 'Date', 'Close']
worksheet1.write_row('A1', myheaders)
worksheet2.write_row('A1', myheaders)
worksheet3.write_row('A1', myheaders)

DTCount1 = 2
Total5Downs = 0
peak = 0
maxdown = 0
recovered = True
maxperdown = 1
i = 1
weeks = 0

while i < sheets.shape[0]:
    if (sheets['Date'][i] - sheets['Date'][i - 1]).days >= 3:
        weeks += 1
    i += 1
i = 0
while i < sheets.shape[0]:
    #check if this is the first row
    if peak == 0:
        peak = sheets['Close'][0]

    #check if the current value is greater than our peak  
    if peak <= sheets['Close'][i]:
        peak = sheets['Close'][i]
        maxdown = sheets['Close'][i]
        my5per = peak * .05
        DTCount1 = Add_to_Worksheet(worksheet1, sheets, i, DTCount1, 'Peak')
    #check if its lower than the downturn we've had
    
    elif sheets['Close'][i] <= maxdown:
        maxdown = sheets['Close'][i]
        while sheets['Close'][i] <= peak - (maxperdown * my5per) :
            
            Total5Downs +=1
            DTCount1 = Add_to_Worksheet(worksheet1, sheets, i, DTCount1, str(maxperdown * 5) + '%')
            recovered = False
            maxperdown += 1
        
    #check if current vlaue recovered half of the downturn
    elif peak - sheets['Close'][i] <= (peak - maxdown)/2 and recovered == False:
        recovered = True
        peak = sheets['Close'][i]
        maxdown = peak
        my5per = peak * .05
        DTCount1 = Add_to_Worksheet(worksheet1, sheets, i, DTCount1, 'Recovered')
        maxperdown = 1
    i += 1

writer.close()
print(weeks)
print(Total5Downs)
print(weeks/Total5Downs)

