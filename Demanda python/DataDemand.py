from readxlxs import xlxstocsv
import csv
import datetime
import openpyxl
   
importedfile = openpyxl.load_workbook(filename = 'DemandaZonas.xlsx', read_only = True, keep_vba = False)
tabnames = importedfile.sheetnames 

###############################################################################
def is_empty(any_structure):
    if any_structure:
        return False
    else:
        return True
       
###########################################################################
substring = "DemandaMensual"
xlxstocsv(tabnames,substring,importedfile)

with open('temp/DemandaMensual.csv') as csvfile:
    readCSV = csv.reader(csvfile, delimiter=',')
    dmdData = [[] for x in range(4)];  
    for row in readCSV:
        if is_empty(row) is True: break 
        for col in range(4):
            val = row[col]
            try: 
                val = float(row[col])
            except ValueError:
                pass
            dmdData[col].append(val)
for col in range(4):
    dmdData[col].pop(0) 

###########################################################################
substring = "Reservas"
xlxstocsv(tabnames,substring,importedfile)

with open('temp/Reservas.csv') as csvfile:
    readCSV = csv.reader(csvfile, delimiter=',')
    dmdDataR = [];  
    for row in readCSV:
        if is_empty(row) is True: break 
        for col in range(1):
            val = row[col]
            try: 
                val = float(row[col])
            except ValueError:
                pass
            dmdDataR.append(val)
    
###########################################################################

substring = "Horario"
xlxstocsv(tabnames,substring,importedfile)
             
with open('temp/Horario.csv') as csvfile:
    readCSV = csv.reader(csvfile, delimiter=',')
    for row in readCSV: columns = len(row); break                
with open('temp/Horario.csv') as csvfile:
    readCSV = csv.reader(csvfile, delimiter=',')
    horaData = [[] for x in range(columns-1)];  
    for row in readCSV:
        if is_empty(row) is True: break 
        for col in range(columns-1):
            val = row[col+1]
            try: 
                val = float(row[col+1])
            except ValueError:
                pass
            horaData[col].append(val)
# for col in range(columns-1):
#     horaData[col].pop(0) 

###########################################################################

substring = "Zonas"
xlxstocsv(tabnames,substring,importedfile)
             
with open('temp/Zonas.csv') as csvfile:
    readCSV = csv.reader(csvfile, delimiter=',')
    for row in readCSV: columns = len(row); break                
with open('temp/Zonas.csv') as csvfile:
    readCSV = csv.reader(csvfile, delimiter=',')
    zonasData = [[] for x in range(columns-1)];  
    for row in readCSV:
        if is_empty(row) is True: break 
        for col in range(columns-1):
            val = row[col+1]
            try: 
                val = float(row[col+1])
            except ValueError:
                pass
            zonasData[col].append(val)
# for col in range(columns-1):
#     zonasData[col].pop(0) 

###########################################################################
month=[[] for x in range(12)]
for y in range(15):
    for m in range(12):
        month[m].append(dmdData[1][12*y+m])
    

# AR example
from statsmodels.tsa.arima.model import ARIMA
for m in range(12):
    for t in range(15):
        # contrived dataset
        data = month[m]
        # fit model
        model = ARIMA(data, order=(1, 1, 1))
        model_fit = model.fit()
        # make prediction
        yhat = model_fit.predict(len(data), len(data), typ='levels')
        month[m].append(yhat[0])

#Mese
q=[[sum(x)/3 for x in zip(month[0], month[1], month[2])],
   [sum(x)/3 for x in zip(month[3], month[10], month[11])],
   [sum(x)/3 for x in zip(month[7], month[8], month[9])],
   [sum(x)/3 for x in zip(month[4], month[5], month[6])]]

# Q1	Ene, Feb, Mar
# Q2	Abr, Dic
# Q3	Ago, Sept, Oct, Nov
# Q4	May, Jun, Jul

promm=[[] for x in range(4)]
for qfila in range(4):
    for y in range(6):
        val=[]
        for m in range(5):
            val.append(q[qfila][5*y+m])
        promm[qfila].append(max(val))

prommlist=[]
for y in range(6):
    for qf in range(4):
        prommlist.append(promm[qf][y])

zondmq=[[] for x in range(15)]
for z in range(15):
    for y in range(24):
        zondmq[z].append(zonasData[z][y+1]*prommlist[y])

final=[[] for x in range(3)]
for z in range(15):
    val=0
    for y in range(24):
        for x in range(24):
            final[0].append(horaData[z+1][0])
            final[1].append(val+1)
            val=val+1
            final[2].append(zondmq[z][y]*horaData[z+1][x+1])


from openpyxl import Workbook

# Create file
wb = Workbook()
dest_filename = 'colombia6/datos_python.xlsx'

ws0 = wb.active
ws0.title = "Summary"

ws0['A2'] = '24 horas';

###########################################################################

# write file results
ws1 = wb.create_sheet(title="Loads")
for k in range(len(final[0])):

    _ = ws1.cell(column=1, row=k+1, value=final[0][k])
    _ = ws1.cell(column=2, row=k+1, value=final[1][k])
    _ = ws1.cell(column=3, row=k+1, value=final[2][k])

###########################################################################

# write file results
ws1 = wb.create_sheet(title="Reservas")
for x in range(len(dmdDataR)):
    for k in range(576):

        _ = ws1.cell(column=1, row=(576*x)+k+1, value=dmdDataR[x])
        _ = ws1.cell(column=2, row=(576*x)+k+1, value=k+1)
        _ = ws1.cell(column=3, row=(576*x)+k+1, value=1)
    
# with open("loads.csv", "w") as f:
#     writer = csv.writer(f)
#     writer.writerows(final)
    
wb.save(filename = dest_filename)