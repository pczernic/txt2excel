import xlsxwriter as x

# tekstowy
read = open('szczcyt-lato-2025.txt', 'r')
data = read.readlines()
read.close()

final = []

for i in range(len(data) - 1):
    data[i] = data[i].replace('\n', '')

    if data[i][0:2] == '  ':
        data[i] = data[i].replace(' ', ',')
        data[i] = data[i].replace(',,,,,,,', ',')
        data[i] = data[i].replace(',,,,,,', ',')
        data[i] = data[i].replace(',,,,,', ',')
        data[i] = data[i].replace(',,,,', ',')
        data[i] = data[i].replace(',,,', ',')
        data[i] = data[i].replace(',,', ',')
        data[i] = data[i][1:]
        final.append(data[i])

    elif data[i][0:2] == '+ ':
        data[i] = data[i].replace('+ ', '')
        data[i] = data[i].replace(' ', ',')
        data[i] = data[i].replace(',,,,,,,', ',')
        data[i] = data[i].replace(',,,,,,', ',')
        data[i] = data[i].replace(',,,,,', ',')
        data[i] = data[i].replace(',,,,', ',')
        data[i] = data[i].replace(',,,', ',')
        data[i] = data[i].replace(',,', ',')
        data[i] = data[i][1:]
        final.append(data[i])

    elif data[i][0:20] == 'Stacje bez zasilania':
        temp = data[i][20:].replace(' ', ',')
        temp = temp.replace(',,,,,,,', ',')
        temp = temp.replace(',,,,,,', ',')
        temp = temp.replace(',,,,,', ',')
        temp = temp.replace(',,,,', ',')
        temp = temp.replace(',,,', ',')
        temp = temp.replace(',,', ',')
        final.append(data[i][0:20] + temp)
    else:
        final.append(data[i])

last_line = data[-1].replace(',', ';')

final.append(last_line[:-1])

for i in range(len(final)):
    final[i] = final[i].split(',')
    if len(final[i]) == 11:
        final[i].pop(-1)
    elif len(final[i]) == 12:
        final[i].pop(-1)
        final[i].pop(-1)

while final[1][0][:3] != 'A.0':
    final.pop(1)

# excel
workbook = x.Workbook('test.xlsx')
worksheet = workbook.add_worksheet()
# formatting
font = workbook.add_format()
font.set_font_name('Times New Roman')
font.set_font_size(10)
font.set_border(1)

centering = font
centering.set_align('center')

wrapping = font
wrapping.set_text_wrap()

border = workbook.add_format()
border.set_border(1)

merge_row = []

for i in range(len(final)):
    if len(final[i]) == 1:
        temp_str = f'B{i + 2}:K{i + 2}'
        worksheet.merge_range(temp_str, '', border)
        worksheet.write_row(i + 1, 1, final[i], centering)
        merge_row = []

    elif len(final[i]) == 2:
        final[i] = [final[i][0] + final[i][1]]
        temp_str = f'B{i + 2}:K{i + 2}'
        worksheet.merge_range(temp_str, '', border)
        worksheet.write_row(i + 1, 1, final[i], centering)
        merge_row = []

    elif len(final[i]) == 5:
        merge_row.append(i)
        temp_str = f'B{i + 2}:K{i + 2}'
        worksheet.write_row(i + 1, 1, final[i], centering)

    else:
        worksheet.write_row(i + 1, 1, final[i], wrapping)
        merge_row = []

# merge_box = f'G{merge_row[0] + 2}:K{merge_row[-1] + 2}'
# worksheet.merge_range(merge_box, '', border)

workbook.close()

# for item in final:
#     print(len(item))