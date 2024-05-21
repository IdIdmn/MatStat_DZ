import openpyxl
import math
import itertools
import string
import scipy

# ----------------------------------- Сюда вписывать свои значения -----------------------------------------

FILENAME = "ТВиМС.xlsx"
IS_LEFT_BORDER_INCLUDED = False # Включена ли левая граница в интервалах: (] - False, [) - True
intervalBorders = [125, 135, 137, 143, 149, 155, 161, 166, 173] # Границы интервалов 

# -----------------------------------------------------------------------------------------------------------

def readValues(height, length):
    allValues = []
    for i in range(1,height + 1):
        for j in range(length):
            num = sheet[columnNames[j] + str(i)].value
            if (type(num) == int or type(num) == float): allValues.append(sheet[columnNames[j] + str(i)].value)
    return allValues

def printCommonStatRow():
    sheet[columnNames[0] + "20"].value = "Xᵢ"
    sheet[columnNames[0] + "21"].value = "nᵢ"
    sheet[columnNames[0] + "22"].value = "cᵢ"
    for i in range(0, len(uniqueValues)):
        sheet[columnNames[i + 1] + "20"].value = uniqueValues[i]
        sheet[columnNames[i + 1] + "21"].value = allValues.count(uniqueValues[i])
        sheet[columnNames[i + 1] + "22"].value = sheet[columnNames[i + 1] + "21"].value / n
    return

def printIntervalRowIncludeRightBorder(*borders):
    sheet[columnNames[0] + "30"].value = "Iᵢ"
    sheet[columnNames[0] + "31"].value = "mᵢ"
    sheet[columnNames[0] + "32"].value = "p*ᵢ"
    sheet[columnNames[0] + "33"].value = "x\u0303ᵢ"
    for i in range(1,len(borders)):
        m = sum([allValues.count(num) for num in uniqueValues if borders[i-1] < num <= borders[i]])
        sheet[columnNames[i] + "30"].value = "(" + str(borders[i - 1]) + ", " + str(borders[i]) + "]"
        sheet[columnNames[i] + "31"].value = m
        sheet[columnNames[i] + "32"].value = m / n
        sheet[columnNames[i] + "33"].value = (borders[i] + borders[i - 1]) / 2
    sheet[columnNames[1] + "30"].value = "[" + str(borders[0]) + ", " + str(borders[1]) + "]"
    sheet[columnNames[1] + "31"].value += allValues.count(borders[0])
    sheet[columnNames[1] + "32"].value = sheet[columnNames[1] + "31"].value / n 
    return

def printIntervalRowIncludeLeftBorder(*borders):
    sheet[columnNames[0] + "30"].value = "Iᵢ"
    sheet[columnNames[0] + "31"].value = "mᵢ"
    sheet[columnNames[0] + "32"].value = "p*ᵢ"
    sheet[columnNames[0] + "33"].value = "x\u0303ᵢ"    
    for i in range(1,len(borders)):
        m = sum([allValues.count(num) for num in uniqueValues if borders[i-1] <= num < borders[i]])
        sheet[columnNames[i] + "30"].value = "[" + str(borders[i - 1]) + ", " + str(borders[i]) + ")"
        sheet[columnNames[i] + "31"].value = m
        sheet[columnNames[i] + "32"].value = m / n
        sheet[columnNames[i] + "33"].value = (borders[i] + borders[i - 1]) / 2
    sheet[columnNames[len(borders) - 1] + "30"].value = "[" + str(borders[-2]) + ", " + str(borders[-1]) + "]"
    sheet[columnNames[len(borders) - 1] + "31"].value += allValues.count(borders[-1])
    sheet[columnNames[len(borders) - 1] + "32"].value = sheet[columnNames[len(borders) - 1] + "31"].value / n 
    return

def printIntervalRow(isLeftBorderIncluded, *borders):
    if isLeftBorderIncluded:
        printIntervalRowIncludeLeftBorder(*borders)
    else:
        printIntervalRowIncludeRightBorder(*borders)

def printLaplasFunctionArguments(*borders):
    sheet[columnNames[0] + "40"].value = "xᵢ"
    sheet[columnNames[0] + "41"].value = "xᵢ - m / σ"
    sheet[columnNames[0] + "42"].value = "Ф(xᵢ - m / σ)"
    for i in range(0,len(borders)):
        sheet[columnNames[i + 1] + "40"].value = borders[i]
        sheet[columnNames[i + 1] + "41"].value = (borders[i] - M) / sigma
        sheet[columnNames[i + 1] + "42"].value = scipy.stats.norm.cdf(sheet[columnNames[i + 1] + "41"].value) - 0.5
    return

def printTheoreticalProbabilities(isLeftBorderIncluded, *borders):
    sheet[columnNames[0] + "45"].value = "Iᵢ"
    sheet[columnNames[0] + "46"].value = "pᵢ"   
    for i in range(1,len(borders)):
        sheet[columnNames[i] + "45"].value = "(" + str(borders[i - 1]) + ", " + str(borders[i]) + "]"
        sheet[columnNames[i] + "46"].value = sheet[columnNames[i + 1] + "42"].value - sheet[columnNames[i] + "42"].value
    if isLeftBorderIncluded: sheet[columnNames[len(borders) - 1] + "45"].value = "[" + str(borders[-2]) + ", " + str(borders[-1]) + "]"
    else: sheet[columnNames[len(borders) - 1] + "45"].value = "[" + str(borders[-2]) + ", " + str(borders[-1]) + "]"
    return

file = openpyxl.load_workbook(FILENAME)
sheet = file.active

columnNames = list(string.ascii_uppercase) + sorted(list(map(lambda x: ''.join(x), list(itertools.combinations_with_replacement(string.ascii_uppercase, 2)))))

height = max([row for row in range(1, 20) if (type(sheet[columnNames[0] + str(row)].value) == int or type(sheet[columnNames[0] + str(row)].value) == float)])
width = max([column + 1 for column in range(20) if (type(sheet[columnNames[column] + str(1)].value) == int or type(sheet[columnNames[column] + str(1)].value) == float)])
allValues = readValues(height, width)
uniqueValues = sorted(list(set(allValues)))
n = len(allValues)
amount_of_intervals = int(1 + math.log2(n)) + 1

printCommonStatRow()

printIntervalRow(IS_LEFT_BORDER_INCLUDED, *intervalBorders)

M = sum([sheet[columnNames[i] + "33"].value * sheet[columnNames[i] + "32"].value for i in range(1, len(intervalBorders))])
D = sum([pow((sheet[columnNames[i] + "33"].value - M), 2) * sheet[columnNames[i] + "32"].value for i in range(1, len(intervalBorders))])
sigma = pow(D, 1/2)

sheet[columnNames[0] + "35"].value = "M[X] = "
sheet[columnNames[1] + "35"].value = M
sheet[columnNames[0] + "36"].value = "D[x] = "
sheet[columnNames[1] + "36"].value = D
sheet[columnNames[0] + "37"].value = "σ[X] = "
sheet[columnNames[1] + "37"].value = sigma

printLaplasFunctionArguments(*intervalBorders)

printTheoreticalProbabilities(IS_LEFT_BORDER_INCLUDED, *intervalBorders)

X = sum([pow(sheet[columnNames[i] + "32"].value - sheet[columnNames[i] + "46"].value, 2) / sheet[columnNames[i] + "46"].value for i in range(1, len(intervalBorders))]) * n
sheet[columnNames[0] + "50"].value = "χ² = "
sheet[columnNames[1] + "50"].value = X

file.save(FILENAME)