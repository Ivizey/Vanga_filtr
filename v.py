import xlrd
import math
import numpy
#розбиваємо масив на необхідну кількість елементів
def split(arr, size):
     arrs = []
     while len(arr) > size:
        pice = arr[:size]
        arrs.append(pice)
        arr = arr[size:]
     arrs.append(arr)
     return arrs
#вставляємо необхідні елементи для створення рівняння
def insertItems(arrForInsertion, q):
    for i in range(0, len(arrForInsertion)+8, 5):
        arrForInsertion.insert(i+1, 1)
        if (q == 0):
            arrForInsertion.insert(i+4, int(arrForInsertion[i+1])*int(arrForInsertion[i+2])*int(arrForInsertion[i+3]))
        elif (q == 1):
            arrForInsertion.insert(i+4, int(arrForInsertion[i+1])*int(arrForInsertion[i])*int(arrForInsertion[i+3]))
        else:
            arrForInsertion.insert(i+4, int(arrForInsertion[i+1])*int(arrForInsertion[i])*int(arrForInsertion[i+3]))
    return arrForInsertion
#формуємо рівняння
def takeTheSystem(data):
    value = []
    for i in range(len(parametr)-2):
        value.append(parametr[i+2])
        value.append(parametr[i])
        value.append(parametr[i+1])
    return value
#нормацізація рівнянь за методом Гауса
def normalize(matrix, k, w1):
    completeEquation = insertItems(matrix, w1)
    mas = split(completeEquation,5)
    row = []
    newMas = []
    for i in range(6):
        for j in range(5):
            row.append(int(mas[i][j])*int(mas[i][int(k)]))
    newMas = (split(row,5))
    for i in range(6):
        print(newMas[i])
    print('---------------Сума----------------')
    suma = [sum(x) for x in zip(*split([int(item) for item in row], 5))]
    print(suma)
    print('-----------------------------------')
    return suma
#5.Обраховуємо середнє арифметичне відхилення
def findDeviations(resData, v):
    print('Модель має вигляд:')
    print(functions[v])
    if (v == 0):
        s = int(resData[0])+int(resData[1])+int(resData[2])+(int(resData[1])*int(resData[2]))
    elif (v == 1):
        s = int(resData[0])+int(resData[1])+int(resData[2])-(int(resData[0])*int(resData[2]))
    else:
        s = int(resData[0])-int(resData[1])-int(resData[2])+(int(resData[0])*int(resData[2]))
    print('F'+str(v+1) ,'=', s)
    print('Середнє квадратичне відхилення:')
    S = 0
    for i in range(len(experemental)):
        S = math.sqrt(pow((int(experemental[v])-(s)),2))
    deviation = (S/len(experemental))
    print("%.3f" % (deviation))
    print(math.sum(Fsum))
#4.знаходимо невідомі коефіцієнти при змінних
def findTheOdds(odds, w2):
    print("Значення коефіцієнтів розв'язаної системи:")
    vector = []
    for i in range(4):
        for j in range(0,5,5):
            vector.append(odds[i][j])
            del odds[i][j]
    result = []
    M2 = numpy.array(odds)
    v2 = numpy.array(vector)
    result = numpy.linalg.solve(M2, v2)
    for i in range(4):
        print('a'+str(i+1),'=', "%.3f" % (result[i]))
    findDeviations(result, w2)
#3.виводимо матриці на екран
def getTheMatrix(inputParametr, w):
    print('Для функції: ', functions[w])
    print('Отримані матриці:')
    equations = []
    for i in range(4):
        equations.append(normalize(takeTheSystem(inputParametr), i+1, w))
    print('Система нормалізованих рівнянь має вигляд:')
    for i in range(4):
        print(equations[i])
    findTheOdds(equations, w)
#1.вибираємо дані з файлу Excel
excel_data_file = xlrd.open_workbook('C:/Users/Павел/xl.xlsx')
sheet = excel_data_file.sheet_by_index(0)
parametr = []
row_number = sheet.nrows
if row_number > 0:
    for row in range(0, row_number):
        parametr.append(str(sheet.row(row)[0]).replace("number", "").replace(":","").replace(".0",""))
    print('Задані початкові умови: ', len(parametr))
else:
    print('Вказаний файл пустий, або дані вказні неправильно!')
#2.викликаємо необхідний метод для розв'язку
functions = ['F1 = a0 + a1 + a2 + a1*a2', 'F2 = a0 + a1 + a2 - a0*a2', 'F3 = a0 - a1 - a2 + a0*a1']
experemental = [13, 14, 15]
print(parametr)
for i in range(len(experemental)):
    getTheMatrix(parametr, i)
    print('                 ===========')
