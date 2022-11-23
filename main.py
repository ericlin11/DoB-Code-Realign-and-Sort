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


# #Sorting the spreadsheet by Section Number
# print("SORTING ...")
# for x in range(2, ws.max_row+1):
#     if ws.cell(x, column=1).value and ws.cell(x, column=3).value is not None:
#         print("same")
#         print("_______________")
#     elif ws.cell(x, column=1).value is not None:
#         nyc = ws.cell(x, column=1).value.split(" ", 1)
#         nycnumber = nyc[0]
#         print(nycnumber)
#         print("_______________")
#     elif ws.cell(x, column=3).value is not None:
#         icc = ws.cell(x, column=3).value.split(" ", 1)
#         iccnumber = icc[0]
#         print(iccnumber)
#         print("_______________")
#
# # sorted(inputlist, key=lambda v: [int(i) for i in v.split('.')])

wb.save('new.xlsx')
print("Done")