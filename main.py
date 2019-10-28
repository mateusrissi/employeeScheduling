# -*- coding: utf-8 -*-

import xlwt 
from xlwt import Workbook
from random import shuffle

onDutyList = [line.strip() for line in open("lists/onDuty.txt")] #on duty people list
mShiftList = [line.strip() for line in open("lists/morningShift.txt")] #morning shift people list
aShiftList = [line.strip() for line in open("lists/afternoonShift.txt")] #afternoon shift people list

mS = len(mShiftList) #morning shift count
aS = len(aShiftList) #afternoon shift count
dC = len(onDutyList) #on duty count

# receive shift list as parameter and return a matrix (nested list)
def matrixGenerator(sL):
    matrix = []
    shiftC = len(sL) #shift list count
    for j in range(dC):
        auxL = [] #auxiliar list
        sCopy = sL[:] #shift list copy
        shuffle(sCopy)
        
        # check if shift list contains the person on duty
        if onDutyList[j] in sCopy:
            sCopy.remove(onDutyList[j])
        
        # append v1, v2 and v3 to auxL
        for _ in range(shiftC-1):
            auxL.append(sCopy.pop())
        
        # append the auxL to the matrix
        matrix.append(auxL)
    return matrix

#calls
matrixM = matrixGenerator(mShiftList)
matrixA = matrixGenerator(aShiftList)

# write to txt file "result.txt"
f = open("result.txt","w+")
f.write("MANHA\n")
f.write("v1 | v2 | v3\n")
for i in range(dC):
    for j in range(mS-1):
        f.write(matrixM[i][j] + " | ")
    f.write("\n")

f.write("\n")

f.write("TARDE\n")
f.write("v1 | v2 | v3\n")
for i in range(dC):
    for j in range(aS-1):
        f.write(matrixA[i][j] + " | ")
    f.write("\n")
f.close()

# write to excel file "result.xls"
wb = Workbook()
sheet1 = wb.add_sheet('Sheet 1')

style0 = xlwt.easyxf('font: name Times New Roman, height 240, bold on; align: vert center, horiz center')
style1 = xlwt.easyxf('font: name Times New Roman, height 240; align: vert center, horiz center')

for i in range(dC):
    sheet1.write(i+2,0, onDutyList[i], style0)

sheet1.write(0, 3, 'MANHA', style0)

sheet1.write(1, 2, 'V1', style0)
sheet1.write(1, 3, 'V2', style0)
sheet1.write(1, 4, 'V3', style0)

for i in range(dC):
    for j in range(mS-1):
        sheet1.write(i+2, j+2, matrixM[i][j], style1)

sheet1.write(0, 8, 'TARDE', style0)

sheet1.write(1, 7, 'V1', style0)
sheet1.write(1, 8, 'V2', style0)
sheet1.write(1, 9, 'V3', style0)

for i in range(dC):
    for j in range(aS-1):
        sheet1.write(i+2, j+7, matrixA[i][j], style1)

wb.save('result.xls')

print("Success!")
