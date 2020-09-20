#takes in an ESTC permalink, and outputs bibliographical information

import requests, bs4, webbrowser, openpyxl, os

os.chdir('/Users/work/Desktop/1509-1593 split')

#open workbook
# wb = openpyxl.load_workbook('Y1583.xlsx')


### select sheet
for currentsheet in range(1509, 1592):
    sheet = wb.get_sheet_by_name(str(currentsheet))
    print(sheet)

    for row in range(111, 265):
        if sheet['A' + str(row)].value == None:
            break
        else:
            URL = sheet['A' + str(row)].value
            print('working on ' + URL)

            ### navigating from URL to main entry page: getting second URL
            res = requests.get(URL)
            res.raise_for_status()
            ESTCsoup1 = bs4.BeautifulSoup(res.text, "html.parser")
            link = str(ESTCsoup1.select('.td1 a'))
            ###strip link of tags etc###
            URL2 = link[10:-8]

            ### URL2 GOT###

            ##fixing link###
            fixedlink = URL2.replace('&amp;', '&')

            ###creating soup from fixedlink###
            res = requests.get(fixedlink)
            res.raise_for_status()
            ESTCsoup = bs4.BeautifulSoup(res.text, "html.parser")

            cells = ESTCsoup.select('.td1') ###all cells in table as a list. RHS cells are odd numbered indexes###

            ###here we need to turn these cells into a list of the contents as just strings:

            tableContentsList = []
            for i in range(0, len(cells)):
                tableContentsList.append(cells[i].getText())

            ###tableContentsList is now a list of all the text (with a lot of white space)

            ###strip the white space
            for item in range(0, len(tableContentsList)):
                tableContentsList[item] = tableContentsList[item].strip()

                

            ###Checking for items in the list and assigning them values

            if 'ESTC System No.' in tableContentsList:
                ESTCSystemNo = tableContentsList[tableContentsList.index('ESTC System No.') + 1]
            else:
                ESTCSystemNo = 'None'

            if 'ESTC Citation No.' in tableContentsList:
                ESTCCitationsNo = tableContentsList[tableContentsList.index('ESTC Citation No.') + 1]
            else:
                ESTCCitationsNo = 'None'

            if 'Author - personal' in tableContentsList:
                Author = tableContentsList[tableContentsList.index('Author - personal') + 1]
            else:
                Author = 'None'

            if 'Uniform title' in tableContentsList:
                UniformTitle = tableContentsList[tableContentsList.index('Uniform title') + 1]
            else:
                UniformTitle = 'None'

            if 'Title' in tableContentsList:
                Title = tableContentsList[tableContentsList.index('Title') + 1]
            else:
                Title = 'None'

            if 'Publisher/year' in tableContentsList:
                Colophon = tableContentsList[tableContentsList.index('Publisher/year') + 1]
            else:
                Colophon = 'None' 

            if 'Physical descr.' in tableContentsList:
                PhysicalDescription = tableContentsList[tableContentsList.index('Physical descr.') + 1]
            else:
                PhysicalDescription = 'None' 

            ###grabbing STC number
            for item in tableContentsList:
                    if item.startswith('STC'):
                            STC = item

            ###it would make more sense to store all the above as a dictionary...maybe something for later

            ##sending to sheet
            sheet['B' + str(row)] = STC
            sheet['C' + str(row)] = ESTCCitationsNo
            sheet['D' + str(row)] = UniformTitle
            sheet['E' + str(row)] = Title
            sheet['F' + str(row)] = Author
            sheet['G' + str(row)] = Colophon
            sheet['H' + str(row)] = PhysicalDescription
            
            ##copy saved as testsubjectreated.xlsx
            wb.save('Y1583.xlsx')
            print('Done')
