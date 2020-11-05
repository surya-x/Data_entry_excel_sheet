# -*- coding: utf-8 -*-

import sys
from app.utils import *
from config import *


data = []
print("\nDecoding PDF unicode...\nThis will take few minutes")
data = get_text_from_pdf(pdf_path)

try:
    all_data = []

    for i, datum in enumerate(data):
        # logging.info("For index %d in list 'data' " % i)

        row_data = []

        if datum != '':
            try:
                lines = datum.splitlines()
                # TODO : remove this print line
                # print("line: " + str(i + 1))

                column_a = lines[12].split(" ")[2].split(".")
                column_a = "/".join(column_a[1:])

                column_b = datum.split(
                    "110381/")[1].split("-")[1].split("\n")[0]

                add_list = []
                add_start = datum.split("Loonbrief")[1]
                add = add_start.split("EUR\n")[0].split("\n")

                for line in add:
                    if line.isupper() or " bus " in line or "(cid:201)" in line:
                        if "TOTAAL " not in line and "VERTROUWELIJK" not in line and len(line) > 4:
                            if "(cid:201)" in line:
                                line = line.replace("(cid:201)", "É")
                            add_list.append(line)

                column_c = ''
                column_d = ''
                column_e = ''
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

                column_f = datum.split("Rijksregisternr.")[1].split("\n")[2]

                if "Uurloon (Dienst. cheq.)" in datum:
                    temp_uurloon = "Uurloon (Dienst. cheq.)"
                elif "Uurloon" in datum:
                    temp_uurloon = "Uurloon"
                column_g = datum.split(temp_uurloon)[1].split("\n")[2]

                if "deeltijds" in datum:
                    temp_tijids = "deeltijds"
                elif "voltijds" in datum:
                    temp_tijids = "voltijds"

                column_h = datum.split(temp_tijids)[1].split("\n")[
                    0].split("/ ")[1]

                column_i = ''
                if "maaltijdcheques" in datum:
                    column_i = datum.split("maaltijdcheques")[
                        0].split("\n")[-1]

                column_j = ''
                if "economische" in datum:
                    column_j = datum.split("economische")[
                        0].split("\n")[-1].split(":")[0]

                column_k = ''
                if "gewerkte" in datum:
                    column_k = datum.split("gewerkte")[0].split(
                        "\n")[-1].split(":")[0]

                column_l = ''
                if "EUR" in datum:
                    column_l = datum.split("EUR")[-1].split("\n")[16]

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

                # print("-----------------------")

                # print(row_data)
                all_data.append(row_data)

                # TODO : remove this line
                # print("writring done for row num " + str(i + 1))
            except Exception as e:
                print("Error : Format of the pdf is changed\nContact developer")
                print(e)
                # logging.error(e)
                sys.exit()

    # print(all_data)
    write_data(excel_path, excel_format, all_data)


except Exception as e:
    print("Error : Format of the pdf is changed\nContact developer")
    print(e)
    sys.exit()

print("\nTask Completed")


