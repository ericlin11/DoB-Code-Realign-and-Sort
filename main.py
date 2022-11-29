import xlsxwriter
import openpyxl

#2024 NYCC Code Revision -BC.xlsx (file to edit)
#after.xlsx (save changes into after.xlsx)
#before.xlsx (file to compare to after.xlsx)

wb = openpyxl.load_workbook('2024 NYCC Code Revision - BC.xlsx')

# for sheet in wb.sheetnames:
#     ws = wb[sheet]
#     print (ws)

ws = wb['BC3']

iccnumber = ""
icctitle = ""
icctext = ""

nycnumber = ""
nyctitle = ""
nyctext = ""

#Check Column 3 for ICC Section
for x in range(2, ws.max_row+1):
    iccsection = ws.cell(x, column=3)
    #If Cell is not empty: split ICC Section into ICC Number and ICC Title and get the ICC Text
    if iccsection.value is not None:
        b = iccsection.value.split(" ", 1)
        iccnumber = b[0]
        icctitle = b[1]
        # print(str(iccsection) + " " + iccnumber + " ."  )
        icctext = ws.cell(x, column=4).value

        #For Each ICC Number, check each cell in Column 1 for same NYC Number
        for y in range(2, ws.max_row+1):
            nycsection = ws.cell(y, column=1)

            #If cell is not empty: split NYC Section into NYC Number and Title
            if nycsection.value is not None:
                a = nycsection.value.split(" ", 1)
                nycnumber = a[0]
                nyctitle = a[1]

                #Check if the NYC Number is the same as ICC Number and if 3rd column of NYC is empty:
                #If Empty 3rd column is ICC Section and 4th column is ICC Text
                if iccnumber == nycnumber and ws.cell(y, column=3).value is None:
                    ws.cell(y, column=3).value = iccsection.value
                    ws.cell(y, column=4).value = icctext
                    print("Aligning.")
                    print(iccsection.value + "\n" + ws.cell(y, column=3).value)
                    print(icctext + "\n" + ws.cell(y, column=4).value)

                    #Delete previous ICC Section from 3rd column and ICC Text from 4th column after realigning
                    ws.cell(x, column=3).value = None
                    ws.cell(x, column=4).value = None

                    print("_______________")

                #If 3rd column is not empty, check if it is already aligned.
                #If it is already aligned, do nothing (continue with next ICC Number)
                elif iccnumber == nycnumber:
                    m = ws.cell(y, column=3).value.split(" ", 1)
                    if m[0] == iccnumber:
                        print(str(ws.cell(y, column=3)) + " - Already Aligned.")
                        #Delete ICC Section from 3rd column and ICC Text from 4th column if it's in different row
                        if x is not y:
                            ws.cell(x, column=3).value = None
                            ws.cell(x, column=4).value = None
                            print("Deleted duplication at " + str(ws.cell(x, column=3)) + " AND " + str(ws.cell(x, column=4)) + ".")
                        print("_______________")
                        continue

                #If 3rd column is not empty and it is already aligned BUT different ICC and NYC Numbers give ERROR on column 5
                #Print ICC Section to column 6 and ICC Text to column 7 aligned to NYC
                elif iccnumber == nycnumber:
                    print("ERROR: " + nycnumber + " IS NOT " + ws.cell(y, column=3).value)
                    ws.cell(y, column=5).value = "ERROR: DIFFERENT SECTION"
                    ws.cell(y, column=6).value = iccsection.value
                    ws.cell(y, column=7).value = icctext
                    print("_______________")


#Sorting the spreadsheet by Section Number


print("SORTING ...")
inputlist = []
for x in range(2, ws.max_row+1):
    if ws.cell(x, column=1).value and ws.cell(x, column=3).value is not None:
        nyc = ws.cell(x, column=1).value.split(" ", 1)
        icc = ws.cell(x, column=3).value.split(" ", 1)
        nycnumber = nyc[0]
        iccnumber = icc[0]
        if nycnumber == iccnumber:
            sectionnumber = ws.cell(x, column=1).value.split(" ", 1)
            inputlist.append({
                        "sectionnumber": sectionnumber[0],
                        "nycsection": ws.cell(x, column=1).value,
                        "nyctext": ws.cell(x, column=2).value,
                        "iccsection": ws.cell(x, column=3).value,
                        "icctext": ws.cell(x, column=4).value,
                    })
        else: print('error' + ws.cell(x, column=1).value + " - " + ws.cell(x, column=3).value)
    elif ws.cell(x, column=1).value is not None:
        nyc = ws.cell(x, column=1).value.split(" ", 1)
        nycnumber = nyc[0]
        inputlist.append({
            "sectionnumber": nycnumber,
            "nycsection": ws.cell(x, column=1).value,
            "nyctext": ws.cell(x, column=2).value,
            "iccsection": ws.cell(x, column=3).value,
            "icctext": ws.cell(x, column=4).value,
        })
    elif ws.cell(x, column=3).value is not None:
        icc = ws.cell(x, column=3).value.split(" ", 1)
        iccnumber = icc[0]
        inputlist.append({
            "sectionnumber": iccnumber,
            "nycsection": ws.cell(x, column=1).value,
            "nyctext": ws.cell(x, column=2).value,
            "iccsection": ws.cell(x, column=3).value,
            "icctext": ws.cell(x, column=4).value,
        })

sortedlist = sorted(inputlist, key=lambda d: d['sectionnumber'])

startingrow = 2

for x in sortedlist:
    ws.cell(startingrow, column=1).value = x["nycsection"]
    ws.cell(startingrow, column=2).value = x["nyctext"]
    ws.cell(startingrow, column=3).value = x["iccsection"]
    ws.cell(startingrow, column=4).value = x["icctext"]
    print(str(ws.cell(startingrow, column=1)) + "\n" +
          str(x["sectionnumber"]) + ": \n" +
          str(x["nycsection"]) + " - " + str(x["nyctext"]) + "\n " +
          str(x["iccsection"]) + " - " + str(x["icctext"]) + "\n________")
    startingrow += 1



wb.save('sorted.xlsx')
print("Done")

# Another way of sorting and realigning.
# 1. Check if column 1 have value. If it does, get their nycsection, nycnumber, nyctitle from column 1, and nyctext from column 2.
# 2. Check the next row of column 1 for values. If it has value. Get their nycnumber. compare it to the existing nycnumber in 'codes'.
# 3. SORT : If smaller, go in front. If bigger, go in back.
# 4. After, check if column 2 have values. Calculate the iccnumber and compare it to nycnumber in 'codes'.
# 5. If iccnumber and iccnumber matches. Add the values iccsection, iccnumber, icctitle, and icctext. If not, add new 'codes'.
# 6. Paste each 'codes' into excel.
# "codes" : [
#     {
#         "id": "",
#         "nycsection": "",
#         "nycnumber": "",
#         "nyctitle": "",
#         "nyctext": "",
#         "iccsection": "",
#         "iccnumber": "",
#         "icctitle" : "",
#         "icctext": "",
#     }
# ]