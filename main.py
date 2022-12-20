import openpyxl

#2024 NYCC Code Revision -BC.xlsx (file to edit)
#after.xlsx (save changes into after.xlsx)
#before.xlsx (file to compare to after.xlsx)

wb = openpyxl.load_workbook('2024 NYCC Code Revision - BC.xlsx')
sheetnum = 5
ws = ws = wb['BC' + str(sheetnum)]

codelist = []
errorlist = []
tablelink1 = ""
tablelink2 = ""
tablelink3 = ""
tablelink1text = ""
tablelink2text = ""
tablelink3text = ""

def get_links():
    print(str(ws.cell(x, column=1)))
    try:
        global tablelink1
        global tablelink1text
        tablelink1 = ws.cell(x, column=5).hyperlink.target
        print(tablelink1)
        tablelink1text = ws.cell(x, column=5).value
        print(tablelink1text)
    except:
        print("No Table 1")
    try:
        global tablelink2
        global tablelink2text
        tablelink2 = ws.cell(x, column=6).hyperlink.target
        print(tablelink2)
        tablelink2text = ws.cell(x, column=6).value
        print(tablelink2text)
    except:
        print("No Table 2")
    try:
        global tablelink3
        global tablelink3text
        tablelink3 = ws.cell(x, column=7).hyperlink.target
        print(tablelink3)
        tablelink3text = ws.cell(x, column=7).value
        print(tablelink3text)
    except:
        print("No Table 3")


for x in range(2, ws.max_row+1):
    if ws.cell(x, column=1).value is not None:
        nycsection = ws.cell(x,column=1).value
        nyctext = ws.cell(x,column=2).value
        get_links()
        try:
            nyc = nycsection.split(" ", 1)
            nycnumber = nyc[0]
            nyctitle = nyc[1]
            codelist.append({
                "sectionnumber": nycnumber,
                "nycsection": nycsection,
                "nyctext": nyctext,
                "iccsection": "",
                "icctext": "",
                "tablelink1": tablelink1,
                "tablelink1text": tablelink1text,
                "tablelink2": tablelink2,
                "tablelink2text": tablelink2text,
                "tablelink3": tablelink3,
                "tablelink3text": tablelink3text,
            })
            tablelink1 = ""
            tablelink2 = ""
            tablelink3 = ""
            tablelink1text = ""
            tablelink2text = ""
            tablelink3text = ""
        except AttributeError:
            print("Not a NYC Code at : " + str(ws.cell(x,column=1)))
            errorlist.append({
                "nycsection": ws.cell(x, column=1).value,
                "nyctext": ws.cell(x, column=2).value,
            })
        except IndexError:
            print("Not a NYC Code at : " + str(ws.cell(x,column=1)))
            errorlist.append({
                "nycsection": ws.cell(x, column=1).value,
                "nyctext": ws.cell(x, column=2).value,
                "iccsection": ws.cell(x, column=3).value,
                "icctext": ws.cell(x, column=4).value,
            })

for x in range(2, ws.max_row+1):
    if ws.cell(x, column=3).value is not None:
        iccsection = ws.cell(x,column=3).value
        icctext = ws.cell(x,column=4).value
        get_links()
        try:
            icc = iccsection.split(" ", 1)
            iccnumber = icc[0]
            icctitle = icc[1]
            if not any(code["sectionnumber"] == iccnumber for code in codelist):
                codelist.append({
                            "sectionnumber": iccnumber,
                            "nycsection": "",
                            "nyctext": "",
                            "iccsection": iccsection,
                            "icctext": icctext,
                            "tablelink1": tablelink1,
                            "tablelink1text": tablelink1text,
                            "tablelink2": tablelink2,
                            "tablelink2text": tablelink2text,
                            "tablelink3": tablelink3,
                            "tablelink3text": tablelink3text,
                        })
                tablelink1 = ""
                tablelink2 = ""
                tablelink3 = ""
                tablelink1text = ""
                tablelink2text = ""
                tablelink3text = ""
            else:
                for code in codelist:
                    if code["sectionnumber"] == iccnumber:
                        code["iccsection"] = iccsection
                        code["icctext"] = icctext
                        # if code["tablelink1"] is not None:
                        #     code["tablelink2"] = tablelink1
                        #     code["tablelink2text"] = tablelink1text
                        # elif code["tablelink2"] is not None:
                        #     code["tablelink3"] = tablelink1
                        #     code["tablelink3text"] = tablelink3
                        continue
        except AttributeError:
            print("Not a ICC Code at: " + str(ws.cell(x, column=3)))
            errorlist.append({
                "nycsection": ws.cell(x, column=1).value,
                "nyctext": ws.cell(x, column=2).value,
                "iccsection": ws.cell(x, column=3).value,
                "icctext": ws.cell(x, column=4).value,
            })

for x in errorlist:
    print("[NYC]: " + str(x["nycsection"]) + " - " + str(x["nyctext"]) + "\n " +
          "[ICC]: " + str(x["iccsection"]) + " - " + str(x["icctext"]) + "\n")

codelist = sorted(codelist, key=lambda d: d["sectionnumber"])
for x in codelist:
    print(str(x["sectionnumber"]) + ": \n" +
        "[NYC]: " + str(x["nycsection"]) + " - " + str(x["nyctext"]) + "\n " +
        "[ICC]: " + str(x["iccsection"]) + " - " + str(x["icctext"]) + "\n" +
        "Table1: " + str(x["tablelink1text"]) + " - " + str(x["tablelink1"]) + "\n" +
        "Table2: " + str(x["tablelink2text"]) + " - " + str(x["tablelink2"]) + "\n" +
        "Table3: " + str(x["tablelink3text"]) + " - " + str(x["tablelink3"]) + "\n")


print("Done")





# for sheet in wb.sheetnames:
#     ws = wb[sheet]
#     print (ws)

# ws = wb['BC3']
#
# iccnumber = ""
# icctitle = ""
# icctext = ""
#
# nycnumber = ""
# nyctitle = ""
# nyctext = ""
#
#
# inputlist = []
#
# n=4
#
# # for n in range (2,2):
# ws = wb['BC' + str(n)]
# print(ws)
#
# def get_links():
#     try:
#         tablelink1 = ws.cell(x, column=5).hyperlink.target
#     except:
#         print("No Table 1")
#     try:
#         tablelink2 = ws.cell(x, column=6).hyperlink.target
#     except:
#         print("No Table 2")
#     try:
#         tablelink3 = ws.cell(x, column=7).hyperlink.target
#     except:
#         print("No Table 3")
#
#
#
#
# #Check Column 3 for ICC Section
# for x in range(2, ws.max_row+1):
#     iccsection = ws.cell(x, column=3)
#     #If Cell is not empty: split ICC Section into ICC Number and ICC Title and get the ICC Text
#     if iccsection.value is not None:
#         b = iccsection.value.split(" ", 1)
#         iccnumber = b[0]
#         icctitle = b[1]
#         # print(str(iccsection) + " " + iccnumber + " ."  )
#         icctext = ws.cell(x, column=4).value
#
#         #Check if it is already aligned but the NYC Number and ICC Number is different (Error)
#         if ws.cell(x, column=1).value is not None:
#             nycsection = ws.cell(x, column=1)
#             a = nycsection.value.split(" ", 1)
#             nycnumber = a[0]
#             if nycnumber != iccnumber:
#                 print("ERROR: " + nycnumber + " IS NOT " + iccnumber + "\n" +
#                                 "Moved ICC to " + str(ws.cell(ws.max_row+1, column=3)))
#                 ws.cell(x, column=3).value = None
#                 ws.cell(x, column=4).value = None
#                 ws.cell(ws.max_row+1, column=3).value = iccsection.value
#                 ws.cell(ws.max_row+1, column=4).value = icctext
#                 continue
#
#
#
#         #For Each ICC Number, check each cell in Column 1 for same NYC Number
#         for y in range(2, ws.max_row+1):
#             nycsection = ws.cell(y, column=1)
#
#             #If cell is not empty: split NYC Section into NYC Number and Title
#             if nycsection.value is not None:
#                 try:
#                     a = nycsection.value.split(" ", 1)
#                 except AttributeError:
#                     print(AttributeError)
#                     continue
#
#                 nycnumber = a[0]
#                 try:
#                     nyctitle = a[1]
#                 except IndexError:
#                     # get_links()
#                     # inputlist.append({
#                     #     "sectionnumber": nycnumber,
#                     #     "nycsection": ws.cell(x, column=1).value,
#                     #     "nyctext": ws.cell(x, column=2).value,
#                     #     "iccsection": ws.cell(x, column=3).value,
#                     #     "icctext": ws.cell(x, column=4).value,
#                     #     "tablelink1": ws.cell(x, column=5).value,
#                     #     "tablelink2": ws.cell(x, column=6).value,
#                     #     "tablelink3": ws.cell(x, column=7).value,
#                     # })
#                     print(IndexError)
#                     continue
#
#
#                 #Check if the NYC Number is the same as ICC Number and if 3rd column of NYC is empty:
#                 #If Empty 3rd column is ICC Section and 4th column is ICC Text
#                 if iccnumber == nycnumber and ws.cell(y, column=3).value is None:
#                     ws.cell(y, column=3).value = iccsection.value
#                     ws.cell(y, column=4).value = icctext
#                     print("Aligning.")
#                     print(iccsection.value + "\n" + ws.cell(y, column=3).value)
#                     print(icctext + "\n" + ws.cell(y, column=4).value)
#
#                     #Delete previous ICC Section from 3rd column and ICC Text from 4th column after realigning
#                     ws.cell(x, column=3).value = None
#                     ws.cell(x, column=4).value = None
#
#                     print("_______________")
#
#                 #If 3rd column is not empty, check if it is already aligned.
#                 #If it is already aligned, do nothing
#                 elif iccnumber == nycnumber and x is y:
#                     print(str(ws.cell(y, column=3)) + " - Already Aligned.")
#                     if x is not y:
#                         ws.cell(x, column=3).value = None
#                         ws.cell(x, column=4).value = None
#                         print("Deleted duplication at " + str(ws.cell(x, column=3)) + " AND " + str(ws.cell(x, column=4)) + ".")
#                     print("_______________")
#
#
#                 #If 3rd column is not empty and it is already aligned BUT different ICC and NYC Numbers give ERROR on column 5
#                 #Print ICC Section to column 6 and ICC Text to column 7 aligned to NYC
#                 # elif x == y and iccnumber != nycnumber:
#                 #     print("ERROR: " + nycnumber + " IS NOT " + str(ws.cell(y, column=3).value) + "\n" +
#                 #           "Moved ICC to " + str(ws.cell(ws.max_row+1, column=3)))
#                 #     print(str(x) + " - " + str(y))
#                 #     # print("Deleted: " + str(ws.cell(x, column=3)))
#                 #     ws.cell(x, column=3).value = None
#                 #     # print("Deleted: " + str(ws.cell(x, column=4)))
#                 #     ws.cell(x, column=4).value = None
#                 #     # ws.cell(ws.max_row+1, column=3).value = iccsection.value
#                 #     # ws.cell(ws.max_row+1, column=4).value = icctext
#                 #     print("_______________")
#
#
# #Sorting the spreadsheet by Section Number.
# #Add NYC Section, NYC Text, ICC Section, ICCText to inputlist(list of dictionary) then sort them by Section Number (ascending)
# print("SORTING ...")
# for x in range(2, ws.max_row+1):
#     #If NYC and ICC exist in same row, check if they have same section number.
#     #If same section number, add them to inputlist, else throw error
#     if ws.cell(x, column=1).value and ws.cell(x, column=3).value is not None:
#         nyc = ws.cell(x, column=1).value.split(" ", 1)
#         icc = ws.cell(x, column=3).value.split(" ", 1)
#         nycnumber = nyc[0]
#         iccnumber = icc[0]
#         tablelink1 = ""
#         tablelink2 = ""
#         tablelink3 = ""
#         if nycnumber == iccnumber:
#             sectionnumber = ws.cell(x, column=1).value.split(" ", 1)
#             get_links()
#             inputlist.append({
#                 "sectionnumber": sectionnumber[0],
#                 "nycsection": ws.cell(x, column=1).value,
#                 "nyctext": ws.cell(x, column=2).value,
#                 "iccsection": ws.cell(x, column=3).value,
#                 "icctext": ws.cell(x, column=4).value,
#                 "tablelink1": tablelink1,
#                 "tablelink2": tablelink2,
#                 "tablelink3": tablelink3,
#             })
#             print()
#         else: print('error' + ws.cell(x, column=1).value + " - " + ws.cell(x, column=3).value)
#     #If only NYC exist, add it to the inputlist.
#     elif ws.cell(x, column=1).value is not None:
#         print(ws.cell(x, column=1).value)
#         print(type(ws.cell(x, column=1).value))
#         nyc = str(ws.cell(x, column=1).value)
#         nycsection = nyc.split(" ", 1)
#         nycnumber = nyc[0]
#         tablelink1 = ""
#         tablelink2 = ""
#         tablelink3 = ""
#         get_links()
#         inputlist.append({
#             "sectionnumber": nycnumber,
#             "nycsection": ws.cell(x, column=1).value,
#             "nyctext": ws.cell(x, column=2).value,
#             "iccsection": ws.cell(x, column=3).value,
#             "icctext": ws.cell(x, column=4).value,
#             "tablelink1": tablelink1,
#             "tablelink2": tablelink2,
#             "tablelink3": tablelink3,
#         })
#     #If only ICC exist, add it to the inputlist
#     elif ws.cell(x, column=3).value is not None:
#         icc = ws.cell(x, column=3).value.split(" ", 1)
#         iccnumber = icc[0]
#         tablelink1 = ""
#         tablelink2 = ""
#         tablelink3 = ""
#         get_links()
#         inputlist.append({
#             "sectionnumber": iccnumber,
#             "nycsection": ws.cell(x, column=1).value,
#             "nyctext": ws.cell(x, column=2).value,
#             "iccsection": ws.cell(x, column=3).value,
#             "icctext": ws.cell(x, column=4).value,
#             "tablelink1": tablelink1,
#             "tablelink2": tablelink2,
#             "tablelink3": tablelink3,
#         })
#
# #sort the inputlist by their section number and save it to sortedlist
# sortedlist = sorted(inputlist, key=lambda d: d['sectionnumber'])
#
# startingrow = 2
#
# #Replacing the cells with the sorted NYC and ICC sections
# for x in sortedlist:
#     ws.cell(startingrow, column=1).value = x["nycsection"]
#     ws.cell(startingrow, column=2).value = x["nyctext"]
#     ws.cell(startingrow, column=3).value = x["iccsection"]
#     ws.cell(startingrow, column=4).value = x["icctext"]
#     ws.cell(startingrow, column=5).value = x["tablelink1"]
#     ws.cell(startingrow, column=6).value = x["tablelink2"]
#     ws.cell(startingrow, column=7).value = x["tablelink3"]
#     print(str(ws.cell(startingrow, column=1)) + "\n" +
#           str(x["sectionnumber"]) + ": \n" +
#           str("[NYC]: " + str(x["nycsection"])) + " - " + str(x["nyctext"]) + "\n " +
#           str("[ICC]: " + str(x["iccsection"])) + " - " + str(x["icctext"]) + "\n________")
#     startingrow += 1
#
# for x in sortedlist:
#     print(x["sectionnumber"])
#
# #Save file to sorted.xlsx
# wb.save('sorted.xlsx')
# print("Done")
#





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