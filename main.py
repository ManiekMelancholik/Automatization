import xlwt
import xlrd
from xlutils import copy


path = "C:/Users/Quiqhaqru/Desktop/MODEL1/pyXLX.xls"
savePath = "C:/Users/Quiqhaqru/Desktop/MODEL1/New_folder/pyXLX.xls"
readBook = xlrd.open_workbook(path)
readSheet = readBook.sheet_by_index(0)


writeBook = copy.copy(readBook)

tags = [[]]
tag_positions = [[]]
casesX = []
casesM = []
casesY = []

def readRowAndSplit(row, row_range, tags2D=[[]], tag_positions=[[]]):
    count = 0
    row_starting_offset = 3
    temp_var = ''
    false_column_iterator = row_starting_offset

    for i in row_range:
        if readSheet.cell_value(rowx=0, colx=i) == '':
            count += 1

    for i in range(count):
        tags2D.append([])
        # tags2D[i].append(readSheet.cell_value(rowx=0, colx=0))
        # tags2D[i].append(readSheet.cell_value(rowx=0, colx=1))
        # tags2D[i].append(readSheet.cell_value(rowx=0, colx=2))

        tag_positions.append([])

        while True:
            temp_var = readSheet.cell_value(rowx=0, colx=false_column_iterator)
            false_column_iterator += 1

            if temp_var == '':
                break

            tags2D[i].append(temp_var)
            tag_positions[i].append(false_column_iterator-1)

def readCollumn(collumn, coll_range,collection):
    for j in coll_range:
        collection.append(readSheet.cell_value(colx=collumn, rowx=j))

def printColl(coll):
    print(coll[0:len(coll)])

def print2DArray(arr2D):
    for j in range(len(arr2D)):
        print(arr2D[j][0:len(arr2D[j])])

headerStyle = xlwt.XFStyle()
headerFont = xlwt.Font()
# headerFont.family = 'Courier New'
headerFont.bold = True
headerFont.height = 400
headerStyle.font = headerFont

importantStyle = xlwt.XFStyle()
importantFont = xlwt.Font()
# importantFont.family = 'Courier New'
importantFont.bold = True
importantFont.height = 300
importantFont.colour_index = 2
importantStyle.font = importantFont

def checkForImportantValues(tags=[],searching_property_name="p"):
    ret_value = -1
    for tag_pos in range(len(tags)):
        if tags[tag_pos] == searching_property_name:
            ret_value = tag_pos
    return ret_value

def constructSheetCasesColumn(sheet=xlwt.Worksheet, cases=[], column=0, tag="tag", selected=[]):
    selected_cases = []
    fake_row_index = 1
    sheet.write(r=0, c=column, label=tag, style=headerStyle)
    if len(selected) == 0:
        selected_cases = range(len(cases))
    else:
        selected_cases = selected

    # print(len(cases))
    # print(cases[0:len(cases)])
    # print(len(selected_cases))
    for row in selected_cases:
        sheet.write(r=fake_row_index, c=column, label=cases[row], style=headerStyle)
        fake_row_index += 1


def fincImportantIndexesInCollumn(col=[], importance_level=0.05):
    ret_arr=[]
    for pos in range(len(col)):
        if col[pos] <= importance_level:
            ret_arr.append(pos)

    # print(ret_arr[0:len(ret_arr)])
    return  ret_arr

def constructSheetColumn(sheet=xlwt.Worksheet, offset=0, tag="tag", values=[] ,check = False, check_for_value = 0.05, important_override_offsets=[]):

    # print(important_override_offsets[0:len(important_override_offsets)])
    sheet.write(r=0, c=offset, label=tag, style=headerStyle)
    row_fake_index = 1
    false_iterator = 0
    for i in range(len(values)):
        if len(important_override_offsets) > 0:
            if important_override_offsets[false_iterator] == i:
                false_iterator += 1
                sheet.write(row_fake_index, offset, values[i], importantStyle)
                row_fake_index += 1
                if false_iterator >= len(important_override_offsets):
                    break

        elif check:
            if check_for_value > values[i]:
                sheet.write(i + 1, offset, values[i], importantStyle)
            else:
                sheet.write(i + 1, offset, values[i])
        else:
            sheet.write(i + 1, offset, values[i])



def constructXLSSheet(sheet_name="new sheet", workbook=xlwt.Workbook(), tag_offsets=[], tags=[], check_for_p=False, check_value_name="p", p_lvl=0.05, only_important=False, indirect=False):
    worksheet = workbook.add_sheet(sheet_name, True)
    # print(tag_offsets[0:len(tag_offsets)])
    values = []
    col_in_sheet = 3
    p_position = -1
    important =[]

    if indirect:
        left = 0
        right = 0
        for tag_pos in range(len(tags)):
            if tags[tag_pos] == "BootLLCI":
                left = tag_pos
            if tags[tag_pos] == "BootULCI":
                right = tag_pos
        # print(readSheet.cell_value(rowx=0, colx=tag_offsets[left]))
        # print(readSheet.cell_value(rowx=0, colx=tag_offsets[right]))
        for i in range(1, readSheet.nrows):
            left_value = readSheet.cell_value(rowx=i, colx=tag_offsets[left])
            right_value = readSheet.cell_value(rowx=i, colx=tag_offsets[right])
            # print(readSheet.cell_value(rowx=i, colx=tag_offsets[left]))
            # print(readSheet.cell_value(rowx=i, colx=tag_offsets[right]))
            if left_value > 0 and right_value > 0:
                important.append(i-1)
                print("+++")
            elif left_value < 0 and right_value < 0:
                important.append(i-1)
                print("---")

        constructSheetCasesColumn(sheet=worksheet, cases=casesX, column=0, tag="X", selected=important)
        constructSheetCasesColumn(sheet=worksheet, cases=casesM, column=1, tag="M", selected=important)
        constructSheetCasesColumn(sheet=worksheet, cases=casesY, column=2, tag="Y", selected=important)

        for i in range(len(tag_offsets)):
            for j in range(1, readSheet.nrows):
                values.append(readSheet.cell_value(rowx=j, colx=tag_offsets[i]))

            constructSheetColumn(sheet=worksheet,
                                 offset=col_in_sheet,
                                 tag=tags[i],
                                 values=values,
                                 important_override_offsets=important
                                 )
            col_in_sheet += 1
            values.clear()

        important.clear()
        return




    if check_for_p or only_important:
        p_position = checkForImportantValues(tags, check_value_name)

    if only_important:
        for j in range(1, readSheet.nrows):
            values.append(readSheet.cell_value(rowx=j, colx=tag_offsets[p_position]))
        # print(readSheet.cell_value(rowx=0, colx=tag_offsets[p_position]))
        # print(values[0:len(values)])
        # print(len(values))
        # print(readSheet.cell_value(rowx=0, colx=p_position))
        important = fincImportantIndexesInCollumn(values)


    values.clear()

    for i in range(len(tag_offsets)):

        for j in range(1, readSheet.nrows):
            values.append(readSheet.cell_value(rowx=j, colx=tag_offsets[i]))

        if only_important:
            constructSheetColumn(sheet=worksheet,
                                 offset=col_in_sheet,
                                 tag=tags[i],
                                 values=values,

                                 important_override_offsets=important
                                 )

        elif i == p_position and p_position != -1:
            constructSheetColumn(sheet=worksheet,
                                 offset=col_in_sheet,
                                 tag=tags[i],
                                 values=values,
                                 check=True
                                 )
        else:
                constructSheetColumn(sheet=worksheet,
                                     offset=col_in_sheet,
                                     tag=tags[i],
                                     values=values
                                     )

        if only_important:
            constructSheetCasesColumn(sheet=worksheet, cases=casesX, column=0, tag="X", selected=important)
            constructSheetCasesColumn(sheet=worksheet, cases=casesM, column=1, tag="M", selected=important)
            constructSheetCasesColumn(sheet=worksheet, cases=casesY, column=2, tag="Y", selected=important)
        else:
            constructSheetCasesColumn(sheet=worksheet, cases=casesX, tag="X", column=0)
            constructSheetCasesColumn(sheet=worksheet, cases=casesM, tag="M", column=1)
            constructSheetCasesColumn(sheet=worksheet, cases=casesY, tag="Y", column=2)

        col_in_sheet += 1
        values.clear()

    important.clear()
    return workbook
def constructWorkbook(sheet_names=[], tag_offsets2D=[[]], tags2D=[[]], workbookToAppend=xlwt.Workbook(), override_sheet_at_index = [], specific="NULL"):
    for i in range(len(sheet_names)):
        if len(override_sheet_at_index) > i:
            if override_sheet_at_index [i]:
                constructXLSSheet(sheet_name=sheet_names[i],
                                  workbook=workbookToAppend,
                                  tag_offsets=tag_offsets2D[i],
                                  tags=tags2D[i],
                                  only_important=True
                                  )
            else:
                constructXLSSheet(sheet_name=sheet_names[i],
                                  workbook=workbookToAppend,
                                  tag_offsets=tag_offsets2D[i],
                                  tags=tags2D[i],
                                  check_for_p=True
                                  )
        elif sheet_names[i] == specific:
            constructXLSSheet(sheet_name=sheet_names[i],
                              workbook=workbookToAppend,
                              tag_offsets=tag_offsets2D[i],
                              tags=tags2D[i],
                              indirect=True
                              )

        else:
            constructXLSSheet(sheet_name=sheet_names[i],
                              workbook=workbookToAppend,
                              tag_offsets=tag_offsets2D[i],
                              tags=tags2D[i]
                              )



active_cases = range(1, readSheet.nrows)

readCollumn(0, active_cases, casesX)
readCollumn(1, active_cases, casesM)
readCollumn(2, active_cases, casesY)
readRowAndSplit(0, range(readSheet.ncols), tags, tag_positions)

writeBook = xlwt.Workbook()

# print2DArray(tags)
# print2DArray(tag_positions)
constructWorkbook(sheet_names=["A", "C", "B", "C+PRIM", "DIRECT", "INDIRECT"],
                  tag_offsets2D=tag_positions,
                  tags2D=tags,
                  workbookToAppend=writeBook,
                  override_sheet_at_index=[False,False,False,False,False],
                  specific="INDIRECT"
                  )

# constructWorkbook(sheet_names=['A','B'],
#                   tag_offsets2D=[[3,4,5,6],[4,5,6,7,8]],
#                   tags2D=[[3,4,5,6],[4,5,6,7,8]],
#                   workbookToAppend=writeBook
#                   )


# writeBook.add_sheet('another', True)

# writeBook.add_sheet("new")
# constructXLSSheet("Sheet A", writeBook, [3,4,5], ['X', 'M', 'Y'])
writeBook.save(savePath)
# printColl(casesX)
# printColl(casesM)
# printColl(casesY)
# print2DArray(tags)
# print2DArray(tag_positions)

