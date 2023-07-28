from tkinter import *
from datetime import date
from utils import get_info_file, get_files_in_folder
import os
import openpyxl.styles
from tkinter.font import Font  # Import the Font class
from openpyxl.cell.text import InlineFont
from openpyxl.cell.rich_text import TextBlock, CellRichText
import time
import glob
class ExcelModifier:
    def __init__(self, file_path, filename, path_to_save):
        self.path_to_save = path_to_save
        self.file_path = file_path
        self.filename = filename
        #print(file_path)
        self.name_file = file_path.split("-")[-1].split(".")[0]

        self.number_of_product, self.number_of_line, self.column_linevf, self.sheet, \
        self.letters, self.name_file, self.workbook, self.version_toprint, self.cpt_update, \
        self.step_data, self.date_value, self.line_date_arrive, self.default_value, self.LANGUAGE, self.step_data2 \
            = get_info_file(file_path)

        self.name_file_U = "U" + str(self.cpt_update) + "-" + str(name_file) + ".xlsx"
        print(self.name_file_U)
        print()

        self.label_list_all = []
        self.buttons_list_all = []
        self.label_date_all = []

        self.flag_first_time_line = True


        self.root = Tk()
        self.root.title(" Create Excel " + self.name_file_U) #+ " - Update " + str(self.cpt_update))
        # Set the font size for the labels to 14
        self.label_font = Font(size=35)
        # Set the font size for the buttons to 12
        self.button_font = Font(size=18)


        self.cpt_general = 0
        self.field_all_var = IntVar(value=1)
        self.field_vars = [StringVar() for _ in range(len(self.step_data) * self.number_of_product + 1)]
        for ix, x in enumerate(self.default_value):
            self.field_vars[ix].set(value=x)
        self.field_vars[-1].set(value="NO")
    def on_next_button_click(self):
        self.root.destroy()

    def modify_excel_multiple(self):
        list_index_1=[]
        update_all_products = self.field_all_var.get()
        if self.LANGUAGE == "GER":
            text_label_pep = "- Peptid "  # + str(peptid+1)
        else:
            text_label_pep = "- Peptide "  # + str(peptid+1)
        #list_index_1.append(text_label_pep)
        for ix, x in  enumerate(self.field_vars):
            peptid = ix // len(self.step_data)
            if update_all_products == 1 and peptid > 0:
                break
            num_id_here = ix % len(self.step_data)
            if ix % len(self.step_data) == 0:
                print(text_label_pep)
                list_index_1.append(text_label_pep+str(peptid+1))
            #offset_line = self.number_of_line + 1
            column_line = self.column_linevf
            selected_value = x.get()
            flag_1more = False
            if selected_value == "NO":
                pass
            elif selected_value == "YES":
                list_index_1.append("*** "+self.step_data[num_id_here][0]+" ***")
            else:
                if selected_value != "Empty":
                    tiret = " -----> "
                else:
                    tiret = ""
                    selected_value = ""
                if selected_value == "ON GOING NPR" or selected_value == "EN COURS NPR":
                    tiret = " -----> "
                    flag_1more = True
                    if self.LANGUAGE == "FR":
                        selected_value = "EN COURS"
                    else:
                        selected_value = "ON GOING"

                list_index_1.append(self.step_data[num_id_here][0] + tiret + selected_value)
                if flag_1more:
                    if self.LANGUAGE == "FR":
                        list_index_1.append("*** NOUVELLE PURIFICATION NECESSAIRE ***")
                    else:
                        list_index_1.append("*** NEW PURIFICATION REQUIRED ***")


        print(list_index_1)
        number_of_line_status_local_start = -1
        for i in range(100):
            if self.sheet[self.letters[self.column_linevf - 1] + str(i + 1)].value == "SUIVI:":
                number_of_line_status_local_start = i + 1
            elif self.sheet[self.letters[self.column_linevf - 1] + str(i + 1)].value == "STATUS:":
                number_of_line_status_local_start = i + 1
                break
        number_of_line_status_local_end = -1
        for i in range(100):
            if self.sheet[self.letters[self.column_linevf - 1] + str(i + 1)].value == "DATE ESTIMEE D'ACHEVEMENT:":
                number_of_line_status_local_end = i + 1
            elif self.sheet[self.letters[self.column_linevf - 1] + str(i + 1)].value == "ANTICIPATED DATE OF COMPLETION :":
                number_of_line_status_local_end = i + 1
                break
        number_of_line = number_of_line_status_local_end - number_of_line_status_local_start - 1
        print(number_of_line_status_local_start, number_of_line_status_local_end)
        if len(list_index_1)+1 == number_of_line:
            pass
        else:
            diff = len(list_index_1)+1 - number_of_line
            for j in range(diff):
                self.sheet.insert_rows(number_of_line_status_local_start + 1)

        number_of_line_status_local_start = -1
        for i in range(100):
            if self.sheet[self.letters[self.column_linevf - 1] + str(i + 1)].value == "SUIVI:":
                number_of_line_status_local_start = i + 1
            elif self.sheet[self.letters[self.column_linevf - 1] + str(i + 1)].value == "STATUS:":
                number_of_line_status_local_start = i + 1
                break
        number_of_line_status_local_end = -1
        for i in range(100):
            if self.sheet[self.letters[self.column_linevf - 1] + str(i + 1)].value == "DATE ESTIMEE D'ACHEVEMENT:":
                number_of_line_status_local_end = i + 1
            elif self.sheet[
                self.letters[self.column_linevf - 1] + str(i + 1)].value == "ANTICIPATED DATE OF COMPLETION :":
                number_of_line_status_local_end = i + 1
                break
        number_of_line = number_of_line_status_local_end - number_of_line_status_local_start - 1
        print(number_of_line_status_local_start, number_of_line_status_local_end)
        offset_line = number_of_line_status_local_start +1
        #ok


        for ix, x in enumerate(list_index_1):
            self.sheet.cell(row=ix + offset_line, column=column_line, value=x)
            if "COMPLETED" in x or "ACHEVEE" in x:
                self.sheet.cell.fill = openpyxl.styles.PatternFill('solid', openpyxl.styles.colors.GREEN)

        self.line_date_arrive = -1
        for i in range(1, 100):
            if self.sheet[self.letters[self.column_linevf - 1] + str(i)].value == "DATE ESTIMEE D'ACHEVEMENT:" or \
                    self.sheet[
                        self.letters[self.column_linevf - 1] + str(i)].value == "ANTICIPATED DATE OF COMPLETION :":
                self.line_date_arrive = i
                break
        if self.line_date_arrive == -1:
            raise "PB"
        else:
            print(" line_date_arrive  ", self.line_date_arrive)
        self.sheet[
            self.letters[self.column_linevf - 1] + str(self.line_date_arrive + 1)].value = self.date_var.get()

        # Save the modified Excel file
        #PATH = f"{self.path_to_save}/{self.name_file}/"
        #if not os.path.exists(PATH):
        #    os.makedirs(PATH)
        modified_file_path = f"{self.path_to_save}/{self.name_file_U}"
        self.workbook.save(modified_file_path)

        print("Excel file modified and saved successfully.")

        os.system(f"open {self.path_to_save}/{self.name_file_U}")

        self.root.destroy()



    def update_radiobuttons_visibility(self):
        update_all_products = self.field_all_var.get()
        if update_all_products == 0:
            self.forget()
            self.start_main_2(num_repeat=self.number_of_product)
        else:
            self.forget()
            self.start_main_2()

    def start_main(self):
        self.field_all_var = IntVar(value=1)
        self.field_all_checkbox = Checkbutton(self.root, text="Update all products with the same status",
                                         variable=self.field_all_var, font=self.button_font, indicatoron=0)
        self.field_all_checkbox.grid(row=0, columnspan=2)
        self.start_main_2()

    def start_main_2(self, num_repeat=1):
        # Create StringVars to store the selected values for Radiobuttons
        self.field_vars = [StringVar() for _ in range(len(self.step_data * self.number_of_product)+1)]
        for i in range(self.number_of_product):
            for ix, x in enumerate(self.default_value):
                self.field_vars[ix+i*len(self.step_data)].set(value=x)

        self.field_vars[-1].set(value="NO")

        self.radiobuttons = []
        self.label0 = Label(self.root, text="  ", font=self.button_font)
        self.label0.grid(row=0 + 2, column=0, sticky=W)


        self.label_list_all = []
        self.buttons_list_all = []
        for numpep in range(num_repeat):
            self.label_list = []
            self.buttons_list = []
            offset = 12 * numpep
            self.label_list.append(Label(self.root, text="  ", font=self.button_font))
            self.label_list[-1].grid(row=0 + 2 + offset, column=0, sticky=W)
            if num_repeat == 1:
                self.label_list.append(Label(self.root, text=" All Peptides " + "--"*10, font=self.button_font))
                self.label_list[-1].grid(row=0 + 3 + offset, column=0, sticky=W)
            else:
                self.label_list.append(Label(self.root, text=" Peptides " + str(numpep + 1) +" "+ "--"*10, font=self.button_font))
                self.label_list[-1].grid(row=0 + 3 + offset, column=0, sticky=W)
            self.label_list.append(
                Label(self.root, text="  ", font=self.button_font))
            self.label_list[-1].grid(row=0 + 4 + offset, column=0, sticky=W)
            for index, (label_text, options) in enumerate(self.step_data):
                #if label_text == "ADD STARS TO THE DATE" and numpep != num_repeat-1:
                #    pass
                #else:
                self.label_list.append(Label(self.root, text=label_text, font=self.button_font))
                self.label_list[-1].grid(row=index + len(self.step_data)+ offset, column=0, sticky=W)
                self.buttons_list.append([])
                for option_index, option in enumerate(options):
                    self.buttons_list[-1].append(Radiobutton(self.root, text=option, variable=self.field_vars[index + numpep * len(self.step_data)],
                                                             value=option, font=self.button_font,
                                                             height=2, width=15, indicatoron=0))
                    self.buttons_list[-1][-1].grid(row=index + len(self.step_data)+ offset, column=option_index + 1, sticky=W)
                    self.radiobuttons.append(self.buttons_list[-1][-1])
                self.label_list_all.append(self.label_list)
                self.buttons_list_all.append(self.buttons_list)

        self.label_list_end = []
        self.buttons_list_end = []
        for index, (label_text, options) in enumerate(self.step_data2):
            self.label_list_end.append(Label(self.root, text=label_text, font=self.button_font))
            self.label_list_end[-1].grid(row=index + len(self.step_data) + offset + 7*num_repeat, column=0, sticky=W)
            self.buttons_list_end.append([])
            for option_index, option in enumerate(options):
                self.buttons_list_end[-1].append(Radiobutton(self.root, text=option, variable=self.field_vars[-1],
                                                         value=option, font=self.button_font,
                                                         height=2, width=15, indicatoron=0))
                self.buttons_list_end[-1][-1].grid(row=index + len(self.step_data) + offset + 7*num_repeat, column=option_index + 1,
                                               sticky=W)
                self.radiobuttons.append(self.buttons_list_end[-1][-1])


        self.label3 = Label(self.root, text="  ", font=self.button_font)
        self.label3.grid(row=len(self.step_data) + len(self.step_data)+ offset, column=0, sticky=W)
        # Add a label and text field for updating the date of receipt
        self.date_label = Label(self.root, text="Date of completion: ", font=self.button_font)
        self.date_label.grid(row=len(self.step_data) + 8+ offset, column=0, sticky=W)

        self.date_var = StringVar()
        self.date_entry = Entry(self.root, textvariable=self.date_var, font=self.button_font)
        self.date_var.set(self.date_value)
        self.date_entry.grid(row=len(self.step_data) + 8+ offset, column=1)

        self.label4 = Label(self.root, text="  ", font=self.button_font)
        self.label4.grid(row=len(self.step_data) + 10+ offset, column=0, sticky=W)

        # Create a button to trigger the modification process

        #if self.field_all_var.get()==1:
        #    self.modify_button = Button(self.root, text="Create Excel", command=self.modify_excel,
        #                                font=self.button_font)
        #else:
        self.modify_button = Button(self.root, text="Create Excel", command=self.modify_excel_multiple,
                                        font=self.button_font)

        self.modify_button.grid(row=len(self.step_data) + 11+ offset, columnspan=len(self.step_data[0][1]) + 1)

        self.next_button = Button(self.root, text="Next", command=self.on_next_button_click, font=self.button_font)
        self.next_button.grid(row=len(self.step_data) + 12+ offset, columnspan=len(self.step_data[0][1]) + 1)


        # Bind the checkbox to the function to update the Radiobuttons visibility
        self.field_all_checkbox.config(command=self.update_radiobuttons_visibility)

        self.root.mainloop()


    def forget(self):
        self.label0.grid_forget()
        del self.label0
        # self.label1.grid_forget()
        # del self.label1
        # self.label2.grid_forget()
        # del self.label2
        self.label3.grid_forget()
        del self.label3
        self.label4.grid_forget()
        del self.label4
        for y in self.label_list_all:
            for x in y:
                x.grid_forget()
                del x
        for z in self.buttons_list_all:
            for x in z:
                for y in x:
                    y.grid_forget()
                    del y
        for x in self.label_date_all:
            x.grid_forget()
            del x
        self.date_label.grid_forget()
        del self.date_label
        # date_var.grid_forget()
        del self.date_var
        self.date_entry.grid_forget()
        del self.date_entry
        self.modify_button.grid_forget()
        del self.modify_button
        self.next_button.grid_forget()
        del self.next_button
        for x in self.buttons_list_end:
            for y in x:
                y.grid_forget()
                del y
        for x in self.label_list_end:
            x.grid_forget()
            del x


    def run(self):
        self.start_main()
        #self.root.mainloop()




if __name__ == "__main__":

    folder_path = "../U0_Updates/"  # Change this to your desired folder path
    files = get_files_in_folder(folder_path)
    timestr = time.strftime("%Y_%m_%d")#_%H_%M_%S")

    PATH = f"../Un_Updates/{timestr}/"
    if not os.path.exists(PATH):
        os.makedirs(PATH)


    txtfiles = []
    for file in glob.glob("../Un_Updates/*/*.xlsx"):
        print(file)
        txtfiles.append(file)
    print(txtfiles)
    #ok

    for filename in files:
        file_path = folder_path + filename
        name_file = file_path.split("-")[-1].split(".")[0]
        #print(name_file)
        #test Un
        update = None
        for i in range(100, 0, -1):
            name_file_U = "U"+str(i)+"-"+str(name_file)+".xlsx"
            #print(name_file_U, txtfiles)
            for x in txtfiles:
                if name_file in x:
                    update = x
        #print(update)
        if update is None:
            #print("U"+str(0)+"-"+str(name_file)+".xlsx")
            excel_modifier = ExcelModifier(file_path,  "U"+str(0)+"-"+str(name_file)+".xlsx", PATH)
            excel_modifier.run()
        else:
            #print(update, name_file_U)
            #ok
            excel_modifier = ExcelModifier(update, name_file_U, PATH)
            excel_modifier.run()
