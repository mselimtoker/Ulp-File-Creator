import pyexcel
from openpyxl import Workbook, load_workbook
import os
from openpyxl.styles import Border, Side
import os
import shutil

a2_reference_table = 'Reference table A2 Engine.xlsx'

u2_reference_table = 'Reference table U2 Engine.xlsx'

temp_start_row = 2

temp_name = 'template.docx'

letter = {'A':0, 'B':1, 'C':2, 'D':3, 'E':4, 'F':5, 'G':6, 'H':7, 'I':8, 'J':9, 'K':10, 'L':11, 'M':12, 'N':13, 'O':14, 'P':15, 'Q': 16, 'R':17, 'S':18, 'T':19, 'U':20, 'V':21, 'W':22, 'X':23, 'Y':24,'Z':25}

def create_file(f_name):
    new_file_name = str(f_name) + ".docx"
    if not os.path.exists("Files"):
        os.makedirs("Files")
    src_dir = os.curdir
    dst_dir = os.path.join(os.curdir, "Files")
    src_file = os.path.join(src_dir, temp_name)
    shutil.copy(src_file, dst_dir)

    dst_file = os.path.join(dst_dir, temp_name)
    new_dst_file_name = os.path.join(dst_dir, new_file_name)
    os.rename(dst_file, new_dst_file_name)


def get_xl(column,row, i, type):
    col = letter[column]
    sr = sheet.row[row]
    if sr[col] == "" or sr[col] == " ":
        print str(row + 1) + ". row has missing value."
    else:
        create_file(sr[col])  # ulp_file


while True:

    while True:
        print "Select your process type. A2/U2: "
        process_type = raw_input()

        if process_type == 'A2' or process_type == 'a2':
            selected_xl = a2_reference_table
            break

        elif process_type == 'U2' or process_type == 'u2':
            selected_xl = u2_reference_table
            break
        else:
            print "Invalid process type !! Please enter A2 or U2."

    if os.path.isfile(selected_xl):
        if os.path.isfile(temp_name):
            book = pyexcel.get_book(file_name=selected_xl)
            is_found = False

            while True:
                print "Please enter tab name for " + selected_xl
                tab_input = raw_input()
                for sheets in book:
                    if tab_input == sheets.name:
                        active_tab_name = tab_input
                        is_found = True

                if is_found:
                    break
                else:
                    print tab_input + " not found. Please check tab name."

            while True:
                try:
                    start_row = int(input("Enter start position of selected row. (Include): "))
                    break
                except:
                    print "You must enter a number"

            while True:
                try:
                    end_row = int(input("Enter end position of selected row. (Include): "))
                    break
                except:
                    print "You must enter a number"
           
            print "Enter column name. Only letter(A or B): "
            col = raw_input()

            sheet = pyexcel.get_sheet(file_name=selected_xl, sheet_name=active_tab_name)

            i = temp_start_row

            for rows in range(start_row - 1, end_row):
                val = get_xl(col, rows, i, process_type)
                i += 1
            print"Process completed."
        else:
            print temp_name + " not found."
    else:
        print selected_xl + " not found."

    while True:
        print "Do you want to repeat process? Y/N"
        repeat = raw_input()

        if repeat == "N" or repeat == "n":
            break
        elif repeat == "Y" or repeat == "y":
            print "Process starts again"
            break
        else:
            print "Invalid entry! Please enter Y or N."

    if repeat == "N" or repeat == "n":
        break
