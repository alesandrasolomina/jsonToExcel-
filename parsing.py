import json
import openpyxl


# Opening JSON file
f = open('result.json')

# returns JSON object as
# a dictionary
data = json.load(f)

name = data['name']
count = 0

# creating a xls file
wb = openpyxl.Workbook()
# # creating a sheet and naming it
sh = wb['Sheet']
sh.title = f'TG messages parsed {name}'

# Iterating through the json
for i in data['messages']:
    text = i['text']

    if text == '':
        continue
    else:
        text = i['text']
        author = i["from"]
        datetime = i["date"]
        datetimeAr = datetime.split("T")
        date = datetimeAr[0]
        time = datetimeAr[1]
        print(author, text)
        count += 1

        for c in range(1, 5):
            cel = sh.cell(count, c)
            if c == 1:
                cel.value = author
            if c == 2:
                cel.value = date
            if c == 3:
                cel.value = time
            if c == 4:
                cel.value = text
        wb.save('/home/aleksandra/Desktop/template.xlsx')
# Closing file
f.close()
