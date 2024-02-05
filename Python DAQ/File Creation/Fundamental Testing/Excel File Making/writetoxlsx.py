import xlsxwriter
import time


def copyarray(data):

    workbook = xlsxwriter.Workbook('output.xlsx')
    worksheet = workbook.add_worksheet()


    row = 0
    col = 0
    for one, two in (data):
        worksheet.write(row, col, one)
        worksheet.write(row, col+1, two)
        row += 1

    worksheet.write(row, col, 'Total')
    worksheet.write(row, col+1, '=SUM(B1:B4)')

    workbook.close()

def timeexcel(n):
    first_time = time.time()
    wb = xlsxwriter.Workbook('Times.xlsx')
    ws = wb.add_worksheet()
    row = 0
    col = 0

    for i in range(n):
        time.sleep(1)
        second_time = time.time()
        ws.write(row, col, i)
        ws.write(row, col+1, first_time)
        ws.write(row, col+2, second_time - first_time)
        row += 1

    wb.close()

def multisheetbook():
    wb = xlsxwriter.Workbook('Double.xlsx')
    ws1 = wb.add_worksheet()
    ws2 = wb.add_worksheet()

    ws1.write(0,0,'This is Worksheet 1')
    ws2.write(0,0,'This is Worksheet 2')

    wb.close()


    
#timeexcel(5)

data = ([0,0],[1,5],[2,10],[3,15])
copyarray(data)

multisheetbook()

