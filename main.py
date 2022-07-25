import xlsxwriter as x

## tekstowy
read = open('szczyt-lato-2020.txt', 'r')
data = read.readlines()
read.close()

print(data[0:7])

final = open('test.txt', 'w')

for i in range(len(data) - 1):

    if data[i][0:2] == '  ':
        data[i] = data[i].replace(' ', ',')
        data[i] = data[i].replace(',,,,,,,', ',')
        data[i] = data[i].replace(',,,,,,', ',')
        data[i] = data[i].replace(',,,,,', ',')
        data[i] = data[i].replace(',,,,', ',')
        data[i] = data[i].replace(',,,', ',')
        data[i] = data[i].replace(',,', ',')
        # print(data[i][0])

        data[i] = data[i][1:]
        # print(data[i])

        final.write(data[i])
    elif data[i][0:20] == 'Stacje bez zasilania':
        temp = data[i][20:].replace(' ', ',')
        temp = temp.replace(',,,,,,,', ',')
        temp = temp.replace(',,,,,,', ',')
        temp = temp.replace(',,,,,', ',')
        temp = temp.replace(',,,,', ',')
        temp = temp.replace(',,,', ',')
        temp = temp.replace(',,', ',')
        final.write(data[i][0:20] + temp)
    else:
        final.write(data[i])

last_line = data[-1].replace(',', ';')
final.write(last_line)
final.close()

## excel
workbook = x.Workbook('test.xlsx')
worksheet = workbook.add_worksheet()

# Widen the first column to make the text clearer.
worksheet.set_column('A:A', 20)

# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})

# Write some simple text.
worksheet.write('A1', 'Hello')

# Text with formatting.
worksheet.write('A2', 'World', bold)

# Write some numbers, with row/column notation.
worksheet.write(2, 0, 123)
worksheet.write(3, 0, 123.456)

workbook.close()