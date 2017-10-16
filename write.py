import xlrd
import cx_Oracle
import glob
import time
import os.path
import datetime
import random, string

# _mtype = 'C:/Users/OchiengB/Desktop/freezie/*.xlsx'
_mtype = 'G:/geek/kenya/*.xl*'

time = time.strftime('%y-%b-%d %H:%M:%S')

Totals1 = 0

rand_str = lambda n: ''.join([random.choice(string.uppercase) for i in xrange(n)])
texted = rand_str(10) 

_mcount = 0
for filename in glob.glob(_mtype):
    

    list = xlrd.open_workbook(filename)
    worksheet = list.sheet_by_index(0)

    f= open("BKR"+texted+".txt","a+")
    NAME = worksheet.cell(2, 1).value
    f.write("FHDR"+NAME+"\n")

    for r in range(4, worksheet.nrows):
        # AGENT_NAME = worksheet.cell(r, 0).value
        # PNR = worksheet.cell(r, 2).value
        CURRENCY_CODE = worksheet.cell(r, 3).value
        rated = worksheet.cell(r, 4).value
       #  PAYMENT_MODE = worksheet.cell(r, 5).value
       #  CURRENCY_CODE = worksheet.cell(r, 6).value
       #  FARE = worksheet.cell(r, 7).value
       #  TAX = worksheet.cell(r, 8).value
       #  FEE = worksheet.cell(r, 9).value
       #  FEE_1 = worksheet.cell(r, 10).value
       #  CREATED_BY = worksheet.cell(r, 11).value

        
        # DB1 = str(Load) + "  " + str(D1) + "  " + str(D2) + "  " + str(D3)+ "  " + str(D4)+ "  " + str(D5)+ "  " + str(D6)+ "  " + str(D7)
        dated = worksheet.cell(3, 4).value
        dated1 = dated.split("_")
        dated2 = datetime.datetime.strptime(dated1[0], "%d-%b-%y")
        START_DATE = dated2.strftime('%d/%m/%Y')
        dated3 = datetime.datetime.strptime(dated1[1], "%d-%b-%y")
        STOP_DATE = dated3.strftime('%d/%m/%Y')
        RATED2 = '{0:016f}'.format(rated).replace('.','')

        Totals1 += rated
        Totals = '{0:016f}'.format(Totals1).replace('.','')
        # rated1 = repr(round(RATED2,4))

        # rated2 rated1.replace('.','')
        CURRENCY_RATES = RATED2

    #     print RATED2

        f.write("DATA"+CURRENCY_CODE+"BKR"+START_DATE+""+STOP_DATE+""+CURRENCY_RATES+"\n")
    columns = str(worksheet.ncols)
    rows = str(worksheet.nrows)
    f.write("FEND"+rows+""+Totals)
    f.close()