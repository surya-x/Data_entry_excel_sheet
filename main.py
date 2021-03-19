# -*- coding: utf-8 -*-

import sys
from app.utils import *
from config import *


data = []
print("\nDecoding PDF unicode...\nThis will take few minutes")
data = get_text_from_pdf(pdf_path)



all_data = []

for i, datum in enumerate(data):
    # logging.info("For index %d in list 'data' " % i)

    row_data = []
    if datum != '':

        lines = datum.splitlines()
        if len(lines) == 0:
            print("Skipping for this file ", i)
            continue

        # TODO : remove this print line
        # print("line: " + str(i + 1))
        # print(lines)
        column_a = ''
        column_b = ''
        column_c = ''
        column_d = ''
        column_e = ''
        column_f = ''
        column_g = ''
        column_h = ''
        column_i = ''
        column_j = ''
        column_k = ''
        column_l = ''
        column_m = ''
        column_n = ''
        column_o = ''

        try:
            column_a = lines[12].split(" ")[2].split(".")
            column_a = "/".join(column_a[1:])
        except Exception as e:
            print("Warning : Something wrong with Column A - Skipping this column")

        try:
            column_b = datum.split("110381/")[1].split("-")[1].split("\n")[0]
        except Exception as e:
            print("Warning : Something wrong with Column B - Skipping this column")

        try:
            add_list = []
            add_start = datum.split("Loonbrief")[1]
            add = add_start.split("EUR\n")[0].split("\n")

            for line in add:
                if line.isupper() or " bus " in line or "(cid:201)" in line:
                    if "TOTAAL " not in line and "VERTROUWELIJK" not in line and len(line) > 4:
                        if "(cid:201)" in line:
                            line = line.replace("(cid:201)", "É")
                        add_list.append(line)

            if len(add_list) == 1:  # only 1 line
                column_c += add_list[0]
            elif len(add_list) == 2:  # only 2 line
                column_c += add_list[0]
                column_d += add_list[1]
            elif len(add_list) == 3:  # only 3 line
                if any(d.isdigit() for d in add_list[-1]):
                    column_e += add_list[-1]
                    column_c += add_list[0]
                    column_d += add_list[1]
                else:
                    column_c += add_list[0] + " " + add_list[1]
                    column_d += add_list[-1]
            elif len(add_list) == 4:  # only 4 line
                if any(d.isdigit() for d in add_list[-1]):
                    column_e += add_list[-1]
                    column_c += add_list[0] + " " + add_list[1]
                    column_d += add_list[2]
                else:
                    column_c += add_list[0] + " " + add_list[1]
                    column_d += add_list[2] + " " + add_list[3]
        except Exception as e:
            print("Warning : Something wrong with Columns C TO E - Skipping the columns")

        try:
            column_f = datum.split("Rijksregisternr.")[1].split("\n")[2]
        except Exception as e:
            print("Warning : Something wrong with Column F - Skipping this column")

        try:
            if "Uurloon (Dienst. cheq.)" in datum:
                temp_uurloon = "Uurloon (Dienst. cheq.)"
            elif "Uurloon" in datum:
                temp_uurloon = "Uurloon"
            column_g = datum.split(temp_uurloon)[1].split("\n")[2]
        except Exception as e:
            print("Warning : Something wrong with Column G - Skipping this column")

        try:
            if "deeltijds" in datum:
                temp_tijids = "deeltijds"
            elif "voltijds" in datum:
                temp_tijids = "voltijds"

            column_h = datum.split(temp_tijids)[1].split("\n")[0].split("/ ")[1]
        except Exception as e:
            print("Warning : Something wrong with Column H - Skipping this column")

        try:
            if "maaltijdcheques" in datum.lower():
                column_i = datum.lower().split("maaltijdcheques")[0].split("\n")[-1]
        except Exception as e:
            print("Warning : Something wrong with Column I - Skipping this column")

        try:
            if "economische" in datum:
                column_j = datum.split("economische")[0].split("\n")[-1].split(":")[0]
        except Exception as e:
            print("Warning : Something wrong with Column J - Skipping this column")

        try:
            if "gewerkte" in datum:
                column_k = datum.split("gewerkte")[0].split("\n")[-1].split(":")[0]
        except Exception as e:
            print("Warning : Something wrong with Column K - Skipping this column")

        try:
            if "werkloosheid overmacht" in datum.lower():
                column_l = datum.lower().split("werkloosheid overmacht")[0].split("\n")[-1].split(" ")[0]
        except Exception as e:
            print("Warning : Something wrong with Column L - Skipping this column")

        try:
            if "uren vakantie" in datum.lower():
                column_m = datum.lower().split("vakantie")[0].split("\n")[-1].split(" ")[0]
        except Exception as e:
            print("Warning : Something wrong with Column M - Skipping this column")

        try:
            if "uren betaalde feestdag" in datum.lower():
                column_n = datum.lower().split("uren betaalde feestdag")[0].split("\n")[-1].split(" ")[0]
        except Exception as e:
            print("Warning : Something wrong with Column N - Skipping this column")

        try:
            if "EUR" in datum:
                column_o = datum.split("EUR")[-1].split("\n")[16]
        except Exception as e:
            print("Warning : Something wrong with Column O - Skipping this column")

        row_data.append(column_a)
        row_data.append(column_b)
        row_data.append(column_c)
        row_data.append(column_d)
        row_data.append(column_e)
        row_data.append(column_f)
        row_data.append(column_g)
        row_data.append(column_h)
        row_data.append(column_i)
        row_data.append(column_j)
        row_data.append(column_k)
        row_data.append(column_l)
        row_data.append(column_m)
        row_data.append(column_n)
        row_data.append(column_o)

        # print(row_data)
        all_data.append(row_data)

        # TODO : remove this line
        print("writring done Page no.  " + str((i*2)+1))
        print("-----------------------")

# print(all_data)
try:
    write_data(excel_path, excel_format, all_data)
    print("\nTask Completed")
except Exception as e:
    print("Error1 : Format of the pdf is changed\nContact developer")
    print(e)
    sys.exit()


