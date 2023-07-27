import openpyxl
import os
from datetime import date
from openpyxl.styles import Color, Fill
from openpyxl.styles import Font


def get_info_file(file_path):
    #name file
    name_file = file_path.split("-")[-1].split(".")[0]
    cpt_update = int(file_path.split("/U")[-1][0])
    flag_first = False
    if cpt_update == 0:
        flag_first = True
    cpt_update += 1
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    #find column
    column_line = 0
    column_linevf = None
    letters =  ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"]
    for i in letters:
        column_line += 1
        if sheet[i + str(1)].value == "Genosphere Biotechnologies":
            column_linevf = column_line
            break
    if column_linevf is None:
        raise "PB"
    else:
        print("Column to look at:", column_linevf, letters[column_linevf - 1])

    # Update date
    current_date = date.today()
    formatted_dateold = sheet[letters[column_linevf - 1] + "5"].value
    if "/" in formatted_dateold:
        formatted_date = current_date.strftime("%d/%m/%Y")
    elif "-" in formatted_dateold:
        formatted_date = current_date.strftime("%d-%m-%Y")
    elif "." in formatted_dateold:
        formatted_date = current_date.strftime("%d.%m.%Y")
    else:
        raise ValueError("Invalid date format in the Excel file.")
    sheet[letters[column_linevf - 1] + "5"].value = formatted_date

    if flag_first:
        for j in range(2):
            for i in range(5, 100):
                if "Peptide" in str(sheet.cell(row=i, column=column_linevf).value) \
                        or "Peptid" in str(sheet.cell(row=i, column=column_linevf).value):
                        #or '"-"&' in str(sheet.cell(row=i, column=column_linevf).value):
                    sheet.delete_rows(i + 1)

    number_of_line = -1
    for i in range(100):
        if sheet[letters[column_linevf-1]+str(i+1)].value == "SUIVI:":
            number_of_line = i + 1
            LANGUAGE = "FR"
        elif sheet[letters[column_linevf - 1] + str(i + 1)].value == "STATUS:":
            number_of_line = i + 1
            LANGUAGE = "EN"
            break
    print(number_of_line)
    number_of_product = min(3, (number_of_line - 12-1))
    if number_of_line ==-1 :
        raise "PB"
    else:
        print("suivi line ", number_of_line)
        print("nombre de produit", number_of_product)
        print(LANGUAGE)

    version_toprint = 0
    for i in range(100):
        if sheet[letters[column_linevf-1]+str(number_of_line+1+i)].value is None :
            version_toprint = i
            break
    if version_toprint ==0 :
        raise "PB"
    else:
        print(" version_toprint ", version_toprint)

    if version_toprint > 4:
        # Create labels and options for each step
        if LANGUAGE == "FR":
            step_data = [
                ("1• Synthèse/déprotection/work up/lyophilisation:", ["EN COURS", "ACHEVEE", "Empty"]),
                ("2• Purification/work up/lyophilisation:-", ["EN COURS", "ACHEVEE", "Empty"]),
                ("3• Modification:-", ["EN COURS", "ACHEVEE", "Empty"]),
                ("4• Purification/work up/lyophilisation:-", ["EN COURS", "ACHEVEE", "Empty"]),
                ("5• Analyses:-", ["EN COURS", "ACHEVEE", "Empty"])
            ]
        else:
            step_data = [
                ("1• Synthesis/deprotection/work up/lyophilisation:", ["ON GOING", "COMPLETE", "Empty"]),
                ("2• Purification/work up/lyophilisation:-", ["ON GOING", "COMPLETE", "Empty"]),
                ("3• Modification:-", ["ON GOING", "COMPLETE", "Empty"]),
                ("4• Purification/work up/lyophilisation:-", ["ON GOING", "COMPLETE", "Empty"]),
                ("5• Analyses:-", ["ON GOING", "COMPLETE", "Empty"])
            ]
    else:
        if LANGUAGE == "FR":
            step_data = [
                ("1• Synthèse/déprotection/work up/lyophilisation:", ["EN COURS", "ACHEVEE", "Empty"]),
                ("2• Purification/work up/lyophilisation:-", ["EN COURS", "ACHEVEE", "Empty"]),
                ("3• Analyses:-", ["EN COURS", "ACHEVEE", "Empty"])
                         ]
        else:
            step_data = [
                ("1• Synthesis/deprotection/work up/lyophilisation:", ["ON GOING", "COMPLETE", "Empty"]),
                ("2• Purification/work up/lyophilisation:-", ["ON GOING", "COMPLETE", "Empty"]),
                ("3• Analyses:-", ["ON GOING", "COMPLETE", "Empty"])
            ]

    line_date_arrive = -1
    for i in range(1, 100):
        if sheet[letters[column_linevf - 1] + str(i)].value == "DATE ESTIMEE D'ACHEVEMENT:" or \
            sheet[letters[column_linevf - 1] + str( i)].value == "ANTICIPATED DATE OF COMPLETION :":
            line_date_arrive = i
            break
    if line_date_arrive ==-1 :
        raise "PB"
    else:
        print(" line_date_arrive  ", line_date_arrive)

    date_value = sheet[letters[column_linevf - 1] + str(line_date_arrive+1)].value
    print(date_value)
    if flag_first:
        sheet.delete_rows(line_date_arrive + 2)
        sheet.delete_rows(line_date_arrive - 2)
    #ok
    print()
    #if flag_first:
    #    ["ON GOING"]
    if flag_first:
        if LANGUAGE == "FR":
            if version_toprint > 4:
                default_value = ["EN COURS", "Empty", "Empty", "Empty", "Empty"]
            else:
                default_value = ["EN COURS", "Empty", "Empty"]
        else:
            if version_toprint > 4:
                default_value = ["ON GOING", "Empty", "Empty", "Empty", "Empty"]
            else:
                default_value = ["ON GOING", "Empty", "Empty"]
    else:
        default_value = []
        flag_goon = False
        if version_toprint > 4:
            num_max = 5
        else:
            num_max = 3
        if LANGUAGE == "FR":
            name_check = "SUIVI:"
        else:
            name_check = "STATUS:"
        if LANGUAGE == "FR":
            for i in range(1, 100):
                #print("OOOOK",sheet[letters[column_linevf - 1] + str(i)].value)
                if flag_goon and len(default_value)<num_max :
                    #print(sheet[letters[column_linevf - 1] + str(i)].value)
                    if str(sheet[letters[column_linevf - 1] + str(i)].value)[-8:] == "EN COURS":
                        default_value.append("EN COURS")
                        #ok
                    elif str(sheet[letters[column_linevf - 1] + str(i)].value)[-7:] == "ACHEVEE":
                        default_value.append("ACHEVEE")
                        #ok
                    else:
                        default_value.append("Empty")
                if sheet[letters[column_linevf - 1] + str(i)].value==name_check:
                    flag_goon = True
                    #print("OK", sheet[letters[column_linevf - 1] + str(i)].value)
        else:
            for i in range(1, 100):
                if flag_goon and len(default_value)<num_max :
                    if str(sheet[letters[column_linevf - 1] + str(i)].value)[-8:] == "ON GOING":
                        default_value.append("ON GOING")
                        #ok
                    elif str(sheet[letters[column_linevf - 1] + str(i)].value)[-7:] == "COMPLETE":
                        default_value.append("COMPLETE")
                        #ok
                    else:
                        default_value.append("Empty")
                if sheet[letters[column_linevf - 1] + str(i)].value==name_check:
                    flag_goon = True



    return number_of_product, number_of_line, column_linevf, sheet, letters, name_file, workbook, version_toprint, \
        cpt_update, step_data, date_value, line_date_arrive, default_value, LANGUAGE

def get_files_in_folder(folder_path):
    return [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]

# Function to modify the Excel file

