import time
import xlsxwriter
start_time = time.time()

# Initial values
length = 1
COM = 0.5
n = 2

counts = 12


# Set up the xlsx file
workbook = xlsxwriter.Workbook('output.xlsx')
worksheet = workbook.add_worksheet()

#Calculate the items
for i in range(counts):
    while length < i:
        length = length + 1 - COM
        COM = ((0.5 * 1) + (1 * n-1)) / (1 + n-1)
        n = n + 1
    timer = time.time()
    worksheet.write(i, 0, timer-start_time)
    worksheet.write(i, 1, n)
    worksheet.write(i, 2, i)

workbook.close()
