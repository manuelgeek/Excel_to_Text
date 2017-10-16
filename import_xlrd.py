import xlrd
import os.path
import glob

_mtype = 'G:/geek/kenya/*.xl*'
for filename in glob.glob(_mtype):
    # fh = open(filename, 'r')

   list = xlrd.open_workbook(filename)
   sh = list.sheet_by_index(0)
   i = 1
   file = open("Output.txt", "w")
   while sh.cell(i,2).value is not 0:
      Load = sh.cell(i,3).value
      D1 = sh.cell(i,3).value
      D2 = sh.cell(i,4).value
      D3 = sh.cell(i,4).value
      D4 = sh.cell(i,3).value
      D5 = sh.cell(i,3).value
      D6 = sh.cell(i,3).value
      D7 = sh.cell(i,3).value
      DB1 = str(Load) + "  " + str(D1) + "  " + str(D2) + "  " + str(D3)+ "  " + str(D4)+ "  " + str(D5)+ "  " + str(D6)+ "  " + str(D7)

      file.write(DB1 + '\n')
      i = i + 1
   file.close()