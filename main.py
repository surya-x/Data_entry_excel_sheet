import sys
from app.utils import *
from config import *

# Path of all the dependenties
pdf_path = r"assets/full_pay.PDF"
excel_path = r"assets/full_pay.xlsx"
excel_format = r"assets/format/excel_format.xlsx"

# write_data(excel_path, excel_format, data=["denis","br"], row_num=2)

try:
    create_date = input(
        "Please Write the created date of pdf in format aa.aa.aaaa \nFor example :- 11.01.2018\n")

    if len(create_date) is not 10:
        print("Enter the correct date format")
        sys.exit()
    elif len(create_date.split(".")) is not 3 or len(create_date.split(".")[2]) is not 4:
        print("Enter the correct date format")
        sys.exit()
except Exception as e:
    print("Enter the correct date format")
    print(e)
    logging.error(e)
    sys.exit()

data = []
print("\nDecoding PDF unicode...\nThis will take few minutes")
data = get_text_from_pdf(pdf_path)

try:
    all_data = []

    for i, datum in enumerate(data):
        logging.info("For index %d in list 'data' " % i)

        row_data = []

        if datum is not '' and (i + 1 <= 102):
            try:
                lines = datum.splitlines()
                # TODO : remove this print line
                print("line: " + str(i + 1))
                logging.info("line: " + str(i + 1))

                column_a = lines[12].split(" ")[2].split(".")
                column_a = "/".join(column_a[1:])
                row_data.append(column_a)

                temp_address = datum.split("VERTROUWELIJK")[1]
                column_b = temp_address.split("\n")[2].split("-")[1]
                row_data.append(column_b)

                try:
                    addr = temp_address.split(create_date)[0].split("\n")[6:-2]
                except Exception as e:
                    print("Error : The date entered is not found")
                    print(e)
                    logging.error(e)
                    sys.exit()

                column_c = ''
                column_d = ''
                column_e = ''

                if len(addr) <= 2:  # only 1 line
                    column_c += addr[0]
                elif len(addr) == 3:  # only 2 line
                    column_c += addr[0]
                    column_d += addr[2]
                elif len(addr) == 5:  # only 3 line
                    if any(d.isdigit() for d in addr[-1]):
                        column_e += addr[-1]
                        column_c += addr[0]
                        column_d += addr[2]
                    else:
                        column_c += addr[0] + " " + addr[2]
                        column_d += addr[-1]
                elif len(addr) == 7:  # only 4 line
                    if any(d.isdigit() for d in addr[-1]):
                        column_e += addr[-1]
                        column_c += addr[0] + " " + addr[2]
                        column_d += addr[4]
                    else:
                        column_c += addr[0] + " " + addr[2]
                        column_d += addr[4] + " " + addr[6]

                row_data.append(column_c)
                row_data.append(column_d)
                row_data.append(column_e)

                column_f = datum.split("Rijksregisternr.")[1].split("\n")[2]
                row_data.append(column_f)

                if "Uurloon (Dienst. cheq.)" in datum:
                    temp_uurloon = "Uurloon (Dienst. cheq.)"
                elif "Uurloon" in datum:
                    temp_uurloon = "Uurloon"

                column_g = datum.split(temp_uurloon)[1].split("\n")[2]
                row_data.append(column_g)

                if "deeltijds" in datum:
                    temp_tijids = "deeltijds"
                elif "voltijds" in datum:
                    temp_tijids = "voltijds"

                column_h = datum.split(temp_tijids)[1].split("\n")[
                    0].split("/ ")[1]
                row_data.append(column_h)

                column_i = ''
                if "maaltijdcheques" in datum:
                    column_i = datum.split("maaltijdcheques")[
                        0].split("\n")[-1]
                row_data.append(column_i)

                column_j = ''
                if "economische" in datum:
                    column_j = datum.split("economische")[
                        0].split("\n")[-1].split(":")[0]
                row_data.append(column_j)

                column_k = ''
                if "gewerkte" in datum:
                    column_k = datum.split("gewerkte")[0].split(
                        "\n")[-1].split(":")[0]
                row_data.append(column_k)

                column_l = ''
                if "EUR" in datum:
                    column_l = datum.split("EUR")[-1].split("\n")[16]
                row_data.append(column_l)

                print(row_data)
                all_data.append(row_data)

                # TODO : remove this line
                print("writring done for row num " + str(i + 1))
            except Exception as e:
                print("Error : Format of the pdf is changed\nContact developer")
                print(e)
                logging.error(e)
                sys.exit()

    # print(all_data)
    write_data(excel_path, excel_format, all_data)


except Exception as e:
    print("Error : Format of the pdf is changed\nContact developer")
    print(e)
    sys.exit()

print("\nTask Completed")
