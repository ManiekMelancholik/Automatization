import xlwt
import xlrd
import xlutils

class Raport:
    def Print_Data(self):
        print("abstr")

    def Add_Data(self, i_data=[""]):
        pass

    def Xls_Write(self, worksheet, offset=0):
        return 0

    class Model:
        def __init__(self, i_lines=[""], name=""):
            self.NAME = name
            tempNames = []
            self.HEADERS = []
            self.CASE_VALUES = []
            # aaa = [[11,21],[32,33]]
            # self.CASE_VALUES.append(aaa)
            self.CASE_TYPES = []
            index = 0
            for l in i_lines:
                split = l.strip("\n").split(" ")

                if len(split) > 1:
                    if len(self.HEADERS) == 0:
                        for i in range(len(split)):
                            if len(split[i]) > 0:
                                # print(split[i])
                                self.HEADERS.append(split[i])

                    elif len(split[0]) > 0:
                        tempArr = []
                        self.CASE_TYPES.append(split[0].rstrip("1"))
                        for n in range(len(tempNames)):
                            tempArr.append(tempNames[n])


                        index += 1
            self.MAX_INDEX = index

            if len(self.CASE_TYPES) == 0:
                self.CASE_TYPES.append("   ")
            # for CT in range(len(self.CASE_TYPES)):
            #     case_arr = []
            #     for H in range(len(self.HEADERS)):
            #         head_arr = []
            #
            #         case_arr.append(head_arr)
            #     self.CASE_VALUES.append(case_arr)

            # if len(self.CASE_TYPES) == 0:
            #     case_arr = []
            #     for H in range(len(self.HEADERS)):
            #         head_arr = [.4, .2]
            #         head_arr.append(0.1)
            #         case_arr.append(head_arr)
            #     self.CASE_VALUES.append(case_arr)

        def PrintHeaders(self):
            print(self.NAME)
            for case in range(len(self.CASE_TYPES)):
                print(self.CASE_TYPES[case])
                # print(self.CASE_VALUES[0][0][0])
                for header in range(len(self.HEADERS)):
                    print(self.HEADERS[header] )
                    print(self.CASE_VALUES[case][header])

        def Add_CaseValues(self, i_lines=[""]):
            if len(self.CASE_VALUES) == 0:
                for CT in range(len(self.CASE_TYPES)):
                    case_arr = []
                    for H in range(len(self.HEADERS)):
                        head_arr = []

                        case_arr.append(head_arr)
                    self.CASE_VALUES.append(case_arr)

            values=[]
            case_type_number = 0
            temp_lines = []
            for l in i_lines:
                values.clear()
                temp_lines.clear()
                split = l.strip("\n").split(" ")
                # print(len(split))
                if len(split) > 0:
                    for splits in split:
                        if len(splits) > 0:
                            temp_lines.append(splits)

                    for i in range(len(temp_lines)):
                        substr = temp_lines[i]
                        if len(substr) > 0:
                            try:
                                float_number = float(substr)
                                values.append(float_number)
                            except:
                               if i == 0:
                                   try:
                                       float(temp_lines[1])
                                       case_type_number += 1
                                   except:
                                       case_type_number -= 1

                    if len(values) > 0:
                        for head in range(len(values)):
                            self.CASE_VALUES[case_type_number][head].append(values[head])

        def Xls_Write(self, worksheet, offset=0):
            style = xlwt.XFStyle()
            font = xlwt.Font()
            font.height = 250
            font.colour_index=2
            style.font = font
            worksheet.write(0, offset+1 ,self.NAME)
            offset += 2

            p_index = []
            llci_index = []

            imp_cases_ind = 0
            if len(self.CASE_TYPES) > 1:
                imp_cases_ind = 1

            for j in range(len(self.CASE_TYPES)):
                for i in range(len(self.HEADERS)):
                    if j-imp_cases_ind >=0:
                        if self.HEADERS[i].strip('\t \n').upper() == "P":
                            p_index.append(i+(j*len(self.HEADERS)))

                        elif self.HEADERS[i].upper().__contains__("LLCI"):
                            llci_index.append(i+(j*len(self.HEADERS)))



                for j in range(len(self.CASE_TYPES)):

                    for i in range(len(self.HEADERS)):
                        # print(len(self.HEADERS))

                        worksheet.write(0, offset + i+(j*len(self.HEADERS)), self.HEADERS[i])

            for case in range(len(self.CASE_TYPES)):
                for header in range(len(self.HEADERS)):
                    if p_index.__contains__(header + (case*len(self.HEADERS))):
                        # if case-imp_cases_ind >= 0:
                            for value in range(len(self.CASE_VALUES[case][header])):
                                p = self.CASE_VALUES[case][header][value]
                                if p<0.05:
                                    worksheet.write(value + 1, offset, p, style)
                                else:
                                    worksheet.write(value + 1, offset, p)
                    elif llci_index.__contains__(header + (case*len(self.HEADERS))):
                        # if case - imp_cases_ind >= 0:
                            for value in range(len(self.CASE_VALUES[case][header])):
                                llci = self.CASE_VALUES[case][header][value]
                                ulci = self.CASE_VALUES[case][header+1][value]
                                if (llci > 0 and ulci > 0) or (llci < 0 and ulci < 0):
                                    worksheet.write(value + 1, offset, llci, style)
                                else:
                                    worksheet.write(value + 1, offset, llci)
                    else:
                        for value in range(len(self.CASE_VALUES[case][header])):
                            worksheet.write(value+1, offset, self.CASE_VALUES[case][header][value])
                    offset += 1

            return offset+1

        def Add_STDcoeff_Header(self, i_data=[""]):
            self.HEADERS.append("STD - " + i_data[0].rstrip("\n \t").upper())

        def Add_STDcoeff_Values(self, i_data=['']):
            for CT in range(1,len(self.CASE_TYPES)):
                for H in range(len(self.HEADERS)):
                    if self.HEADERS[H].__contains__("STD - "):
                        split = i_data[CT].strip("\n\t").split(" ")
                        for S in range(1, len(split)):
                            if len(split[S])>0:
                                self.CASE_VALUES[CT][H].append(split[S])

        def Get_Case_Values_By_Names(self, case_name="",case_number=0, variable_names=[""]):
            if len(case_name) == 0:
                case_name=self.CASE_TYPES[0]

            return_values=[]
            case_index = self.CASE_TYPES.index(case_name)
            for variable in variable_names:
                for H in range(len(self.HEADERS)):
                    if self.HEADERS[H].upper() == variable.upper():
                        return_values.append(self.CASE_VALUES[case_index][H][case_number])


                # return_values.append(self.CASE_VALUES[self.CASE_TYPES[self.CASE_TYPES.index(case_name)]][case_number][ self.HEADERS.index(variable)])
            return return_values


class ModelCases(Raport):
    def __init__(self, i_data=[""]):
        self.HEADERS=[]
        self.CASES=[]
        for l in i_data:
            split = l.strip(" \t \n").split(":")
            if not len(split[0]) == 0 and not split[0].upper().__contains__('SIZE'):
                self.HEADERS.append(split[0])


    def Print_Data(self):
        for l in range(len(self.HEADERS)):
            print(self.HEADERS[l])
            # print(len(self.HEADERS))
            # print(len(self.CASES[0]))
            for i in range(len(self.CASES)):
                print(self.CASES[i][l])


    def Add_Data(self, i_data=[""]):
        case =[]

        for l in i_data:
            split = l.strip(" \t \n").split(":")
            if len(split) > 0 :
                for i in range(1, len(split)):
                    if not len(split[i]) == 0:
                        case.append(split[i])
        case.append("xx")
        self.CASES.append(case)


    def Xls_Write(self, worksheets=[], offset=0):
        for worksheet in worksheets:
            for i in range(len(self.HEADERS)):
                worksheet.write(offset,i, self.HEADERS[i])

            for i in range(len(self.CASES)):
                for j in range(len(self.CASES[i])):
                    worksheet.write(i+1,j+offset,self.CASES[i][j])

        return 4

    def Get_Cases_Amount(self):
        return len(self.CASES)
    def Get_Specific_Case(self, case_index=0):
        return self.CASES[case_index]

class ModelPath(Raport):

    def __init__(self, i_data=[""], i_path_type=""):
        # for l in i_data:
        #     print(l)

        for i in range(len(i_data)):
            if i == len(i_data)-1:
                break
            elif i_data[i].upper().__contains__("MODEL SUMMARY"):
                tempI = i + 1
                while not i_data[tempI].upper().__contains__("MODEL"):
                    tempI += 1
                self.MODEL_SUMMARY = Raport.Model(i_data[i+1:tempI-1],"MODEL SUMMARY")
                i = tempI-2

            elif i_data[i].upper().__contains__("MODEL"):
                tempI = i + 1
                while not i_data[tempI].upper().__contains__("Standardized".upper()):
                    tempI += 1
                self.MODEL = Raport.Model(i_data[i+1:tempI-1],"MODEL")
                i = tempI-2

            elif i_data[i].upper().__contains__("Standardized".upper()):
                tempI = i + 1
                while not i_data[tempI].upper().__contains__("***"):
                    tempI += 1
                self.MODEL.Add_STDcoeff_Header(i_data[i+1:tempI-1])
                i = tempI-2
        # print("\n\ninit done\n\n")

    def Print_Data(self):
        self.MODEL_SUMMARY.PrintHeaders()
        self.MODEL.PrintHeaders()
        # print(self.path_type)
        # for l in self.data:
        #     print(l)

    def Add_Data(self, i_data=[""]):
        for i in range(len(i_data)):
            if i == len(i_data)-1:
                break

            elif i_data[i].upper().__contains__("MODEL SUMMARY"):
                tempI = i + 1
                while not i_data[tempI].upper().__contains__("MODEL"):
                    tempI += 1
                self.MODEL_SUMMARY.Add_CaseValues(i_data[i+1:tempI-1])
                i = tempI-2

            elif i_data[i].upper().__contains__("MODEL"):
                tempI = i + 1
                while not i_data[tempI].upper().__contains__("Standardized".upper()):
                    tempI += 1
                self.MODEL.Add_CaseValues(i_data[i+1:tempI-1])
                i = tempI-2

            elif i_data[i].upper().__contains__("Standardized".upper()):
                tempI = i + 1
                while not i_data[tempI].upper().__contains__("***"):
                    tempI += 1
                # self.STD_COEFF = ModelPath.StdCoeff()
                self.MODEL.Add_STDcoeff_Values(i_data[i + 1:tempI - 1])
                i = tempI-2
        # print("\n\ninit done\n\n")

    def Xls_Write(self, worksheet, offset=0):
        offset = self.MODEL_SUMMARY.Xls_Write(worksheet, offset)
        offset = self.MODEL.Xls_Write(worksheet, offset)
        return offset

    def Get_Case_Values_By_Names(self, case_name="", case_number=0, variable_names=[""]):
        return self.MODEL.Get_Case_Values_By_Names(case_name,case_number,variable_names)

class ModelIndirect(Raport):
    path_type=""
    # models = [Model]

    def __init__(self, i_data=[""], i_path_type=""):


        # for l in i_data:
        #     print(l)

        for i in range(len(i_data)):
            if i == len(i_data)-1:
                break

            elif i_data[i].upper().__contains__("COMPLETELY"):
                tempI = i + 1
                while not i_data[tempI].upper().__contains__("***"):
                    tempI += 1
                self.STANDARDIZED = ModelIndirect.Model(i_data[i+1:tempI-1],"COMPLETELY INDIRECT")
                i = tempI

            elif i_data[i].upper().__contains__("INDIRECT"):
                tempI = i + 1
                while not i_data[tempI].upper().__contains__("COMPLETELY"):
                    tempI += 1
                self.INDIRECT = ModelIndirect.Model(i_data[i+1:tempI-1],"INDIRECT")
                i = tempI


        # print("\n\ninit done\n\n")

    def Print_Data(self):
        self.STANDARDIZED.PrintHeaders()
        self.INDIRECT.PrintHeaders()
        # print(self.path_type)
        # for l in self.data:
        #     print(l)

    def Add_Data(self, i_data=[""]):
        for i in range(len(i_data)):
            if i == len(i_data) - 1:
                break

            elif i_data[i].upper().__contains__("COMPLETELY"):
                tempI = i + 1
                while not i_data[tempI].upper().__contains__("***"):
                    tempI += 1
                self.STANDARDIZED.Add_CaseValues(i_data[i + 1:tempI - 1])
                i = tempI

            elif i_data[i].upper().__contains__("INDIRECT"):
                tempI = i + 1
                while not i_data[tempI].upper().__contains__("COMPLETELY"):
                    tempI += 1
                self.INDIRECT.Add_CaseValues(i_data[i + 1:tempI - 1])
                i = tempI

    def Xls_Write(self, worksheet, offset=0):
        offset = self.STANDARDIZED.Xls_Write(worksheet,offset)
        offset = self.INDIRECT.Xls_Write(worksheet,offset)
        return offset

    def Get_Case_Values_By_Names(self, case_name="", case_number=0, variable_names=[""]):
        if case_name.upper()=="IDIR":
            return self.INDIRECT.Get_Case_Values_By_Names(case_number=case_number,variable_names=variable_names)
        if case_name.upper()=="SIDIR":
            return self.STANDARDIZED.Get_Case_Values_By_Names(case_number=case_number,variable_names=variable_names)


txt_file_path = "C:/Users/Quiqhaqru/Desktop/SpssTxtPython.txt"
test_path = "C:/Users/Quiqhaqru/Desktop/PythoonTesting/TEST_TXT.txt"

xls_file_path = "C:/Users/Quiqhaqru/Desktop/SpssXlsPython.xls"
objects = []


def HeadersFromFile(text=[""]):

    sheet_iterator = 0
    for i in range(len(text)):

#   ---XYM

        if text[i].__contains__("#@#:"):
            return
        elif text[i].__contains__("***") and text[i+1].upper().__contains__("MODEL"):
            _i = i + 1
            while not text[_i].__contains__("***"):
                _i += 1

            objects.append(ModelCases(text[i+2: _i]))

            i += _i

#   ---TOTAL DIRECT INDIRECT COMPSTD

        elif text[i].__contains__("***") and text[i].upper().__contains__("TOTAL, DIRECT,"):
            _i = i + 1
            while not text[_i].__contains__("***"):
                _i += 1

            objects.append(ModelIndirect(text[i+2: _i+1], "TOTAL DIR"))

            i += _i

#   ---C TOTAL EFFECT

        elif text[i].__contains__("***") and text[i].upper().__contains__("TOTAL EFFECT"):
            _i = i + 1
            while not text[_i].__contains__("***"):
                _i += 1

            objects.append(ModelPath(text[i+2: _i+1], "A or B or Cprim"))

            i += _i


#   ---ABC'

        elif text[i].__contains__("***") and text[i+1].upper().__contains__("OUTCOME VARIABLE"):
            _i = i + 1
            while not text[_i].__contains__("***"):
                _i += 1

            objects.append(ModelPath(text[i+2: _i+1], "A or B or Cprim"))

            i += _i


def DataFromFile(text=[""]):

    sheet_iterator = 0
    # print(len(text))
    for i in range(len(text)):

#   ---XYM

        if text[i].__contains__("#@#:"):
            return
        elif text[i].__contains__("***") and text[i+1].upper().__contains__("MODEL"):
            _i = i + 1
            while not text[_i].__contains__("***"):
                _i += 1

            temp_obj = objects[sheet_iterator]
            temp_obj.Add_Data(text[i + 2: _i + 1])
            sheet_iterator += 1
            i += _i

#   ---TOTAL DIRECT INDIRECT COMPSTD

        elif text[i].__contains__("***") and text[i].upper().__contains__("TOTAL, DIRECT,"):
            _i = i + 1
            while not text[_i].__contains__("***"):
                _i += 1

            temp_obj = objects[sheet_iterator]
            temp_obj.Add_Data(text[i + 2: _i + 1])
            sheet_iterator += 1
            i += _i

#   ---C TOTAL EFFECT

        elif text[i].__contains__("***") and text[i].upper().__contains__("TOTAL EFFECT"):
            _i = i + 1
            while not text[_i].__contains__("***"):
                _i += 1
            # print(sheet_iterator)
            temp_obj = objects[sheet_iterator]
            temp_obj.Add_Data(text[i+2: _i+1])
            sheet_iterator += 1
            i += _i


#   ---ABC'

        elif text[i].__contains__("***") and text[i+1].upper().__contains__("OUTCOME VARIABLE"):
            _i = i + 1
            while not text[_i].__contains__("***"):
                _i += 1

            temp_obj = objects[sheet_iterator]
            temp_obj.Add_Data(text[i + 2: _i + 1])
            sheet_iterator += 1

            i += _i


with open(txt_file_path, 'r') as file:
    # for line in file.readlines():
    #     print(line)
    TXT = file.readlines()
    starting_line = 0
    ending_line = 0
    for i in range(len(TXT)):
        if TXT[i].upper().__contains__("RUN MATRIX"):
            starting_line = i
        if TXT[i].__contains__("#@#:"):
            ending_line = i
            break

    HeadersFromFile(TXT[starting_line:ending_line])


    # DataFromFile(TXT[starting_line:ending_line])

    reading = True
    while reading:
        DataFromFile(TXT[starting_line:ending_line])
        if ending_line >= len(TXT):
            reading = False
            break

        for i in range(ending_line+1, len(TXT)):

            if TXT[i].upper().__contains__("RUN MATRIX"):
                starting_line = i
            if TXT[i].__contains__("#@#:"):
                ending_line = i+1
                break


workbook = xlwt.Workbook(xls_file_path)
sheet_names =['A','Cprim & B','C','IDIR & SIDIR']
sheets = []

for name in sheet_names:
    sheets.append(workbook.add_sheet(sheetname=name, cell_overwrite_ok=True))

coll_offset = objects[0].Xls_Write(sheets)

for i in range(len(sheets)):
    objects[i+1].Xls_Write(sheets[i], coll_offset)

workbook.save(xls_file_path)

###objects 0 - all cases
###objects 1 - a
###objects 2 - Cprim & B
###objects 3 - C (Total)
###objects 4 - IDIR & SIDIR


xls_results_path = "C:/Users/Quiqhaqru/Desktop/SpssXlsPython.xls"

ABCCPRIM_CASES = ["std -           coeff", "se", "t", "p","llci","ulci"]
def Cases(case =0):
    return objects[0].Get_Specific_Case(case)
    pass
def A(case=0):
    return objects[1].Get_Case_Values_By_Names(case_name="KPK",case_number=case, variable_names=ABCCPRIM_CASES)
    pass

def Cprim(case=0):
    return objects[2].Get_Case_Values_By_Names(case_name="KPK", case_number=case, variable_names=ABCCPRIM_CASES)
    pass

def B(case=0):
    return objects[2].Get_Case_Values_By_Names(case_name="PCI", case_number=case, variable_names=ABCCPRIM_CASES)
    pass

def C(case=0):
    return objects[3].Get_Case_Values_By_Names(case_name="KPK", case_number=case, variable_names=ABCCPRIM_CASES)
    pass

def IDIR(case=0):
    return objects[4].Get_Case_Values_By_Names(case_name="IDIR", case_number=case, variable_names=["effect", "bootse", "bootllci", "bootulci"])
    pass

def SIDIR(case=0):
    return objects[4].Get_Case_Values_By_Names(case_name="SIDIR", case_number=case, variable_names=["effect", "bootse", "bootllci", "bootulci"])
    pass



analyzes_row_indexes =[0,0,0,0,0]

def ALL(case=0,indexes=[0,0,0,0,0],sheets=[]):
    a=A(case)
    b=B(case)
    c=C(case)# TOTAL
    cp=Cprim(case)#LAST
    case_variable_names = Cases(case)
    # idir=IDIR(case)
    sidir=SIDIR(case)#AxB
    b1=False #sidir
    b2=False #c
    b3=False #cp
    analyze_text = ""
    analyze_ind = 0

    #["coeff", "se", "t", "p","llci","ulci"]
    if (sidir[2] > 0 and sidir[3] > 0) or (sidir[2] < 0 and sidir[3] < 0):
        b1 = True

    if(c[4] > 0 and c[5] > 0) or (c[4] < 0 and c[5] < 0):
        if(c[3] < 0.05):
            b2 = True

    if(cp[4] > 0 and cp[5] > 0) or (cp[4] < 0 and cp[5] < 0):
        if(cp[3] < 0.05):
            b3 = True

    if b1:
        if b2:
            if b3:
                analyze_ind=0
                print(
                    f"""
                    -------------------------
                    {case}
                    {Cases(case)}
                    Complementary
                    MEDIATION
                    
                    {a}
                    {b}
                    {c}
                    {cp}
                    {sidir}
                    -------------------------
                    """)
            else:
                analyze_ind=1
                print(
                    f"""
                    -------------------------
                    {case}
                    {Cases(case)}
                    Competitive
                    MEDIATION
                    
                    {a}
                    {b}
                    {c}
                    {cp}
                    {sidir}
                    -------------------------
                    """)
        else:
            analyze_ind=2
            print(
                f"""
                -------------------------
                {case}
                {Cases(case)}
                Indirect-only
                MEDIATION
                
                {a}
                {b}
                {c}
                {cp}
                {sidir}
                -------------------------
                """)
            analyze_text = f"""
Pomiędzy podanymi skalami dochodzi do mediacji pośredniczącej. Zgodnie z modelem analizowania mediacji (Zhao, Lynch, Chen 2010), można zauważyć istotność ścieżki (SIDIR) pomiędzy X: {case_variable_names[1]}, Y: {case_variable_names[0]}, przy udziale M: {case_variable_names[2]}, co świadczy o pośredniczącym wpływie zmiennej objaśniającej na objaśnianą. Jednocześnie ścieżka (C) jest nieistotna, czyli zmienna X nie wpływa bezpośrednio na zmienną Y. Przedziały ufności dla ścieżki (SIDIR) nie zawierają zera CI:[{sidir[2]}  {sidir[3]}], co dla tego typu efektu świadczy o istotności statystycznej, a standaryzowana beta (STD - BETA) = {sidir[0]}.
"""
    else:
        if b2:
            analyze_ind=3
            # return 0
            print(
                f"""
                -------------------------
                {case}
                {Cases(case)}
                Direct only
                NO MEDIATION

                {a}
                {b}
                {c}
                {cp}
                {sidir}
                -------------------------
                """)
            analyze_text = f"""
Pomiędzy podanymi skalami nie dochodzi do mediacji. Jednakże zgodnie z modelem analizowania mediacji (Zhao, Lynch, Chen 2010), można zauważyć istotność ścieżki (C) pomiędzy {case_variable_names[1]} oraz {case_variable_names[0]}, co świadczy o bezpośrednim wpływie zmiennej objaśniającej na objaśnianą. Przedziały ufności dla ścieżki (C) nie zawierają zera CI:[{c[4]}  {c[5]}], a poziom prawdopodobieństwa (p) = {c[3]} spełnia warunek p < .05. 
"""
        else:
            analyze_ind=4
            # return 0
            print(
                f"""
                -------------------------
                {case}
                No-effect
                NO MEDIATION
                -------------------------
                """)


    # sheets[analyze_ind].write()
    row_index = indexes[analyze_ind]
    col_index = 0
    top_headers = ["STD - BETA", "SE", "t", "p", "95% LLCI", "95% ULCI"]
    left_headers = ["Ścieżka", "A", "B", "C", "C'", "", "SIDIR"]
    bot_headers = ["beta", "SE", "LLCI", "ULCI"]

    sheets[analyze_ind].write(row_index,col_index,
                              f"""
Tabela XX: Ścieżki analizy mediacji dla X:({case_variable_names[1]}), M:({case_variable_names[2]}), Y:({case_variable_names[0]}) przy wielkości próby n = {case_variable_names[3]}.
(STD - BETA) - standaryzowany wynik beta, (SIDIR) - standaryzowany wynik ścieżki pośredniej, (LLCI) - dolny przedział ufności, (ULCI) - górny przedział ufności, przedziały ufności dla 95% 
""")
    row_index += 1
    for j in range(len(left_headers)):
        sheets[analyze_ind].write(row_index, col_index,left_headers[j])
        row_index += 1

    row_index = indexes[analyze_ind] + 1

    for i in range(len(top_headers)):
        col_index += 1
        sheets[analyze_ind].write(row_index,col_index,top_headers[i])

    col_index = 1
    row_index += 1

    for var in a:
        sheets[analyze_ind].write(row_index,col_index,var)
        col_index += 1
    row_index += 1
    col_index = 1

    for var in b:
        sheets[analyze_ind].write(row_index, col_index, var)
        col_index += 1
    row_index += 1
    col_index = 1

    for var in c:
        sheets[analyze_ind].write(row_index, col_index, var)
        col_index += 1
    row_index += 1
    col_index = 1

    for var in cp:
        sheets[analyze_ind].write(row_index, col_index, var)
        col_index += 1
    row_index += 1
    col_index = 1

    for var in bot_headers:
        sheets[analyze_ind].write(row_index,col_index, var)
        col_index += 1
    row_index += 1
    col_index = 1

    for var in sidir:
        sheets[analyze_ind].write(row_index, col_index, var)
        col_index += 1
    row_index += 1
    col_index = 0
    sheets[analyze_ind].write(row_index, col_index,analyze_text)
    row_index += 1
    indexes[analyze_ind] = row_index + 1

    return indexes

# count = 0


analyzes_workbook = xlwt.Workbook()
analyzes_sheets = []
analyzes_names = ["Complementary", "Competitive", "Indirect-Only", "Direct-Only", "No effect"]

for i in range(5):
    analyzes_sheets.append(analyzes_workbook.add_sheet(analyzes_names[i], True))



xls_analyzes_file = "C:/Users/Quiqhaqru/Desktop/Analyzes.xls"


for i in range(objects[0].Get_Cases_Amount()):
    analyzes_row_indexes = ALL(i, analyzes_row_indexes, analyzes_sheets)

analyzes_workbook.save(xls_analyzes_file)