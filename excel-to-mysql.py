import pandas as pd

excel_file = pd.ExcelFile("test.xlsx")  # insert file here
script_file_path = "sql-script.txt"  # results in text file
sheet_lists = excel_file.sheet_names

# results_file = open(script_file_path, "w")
# results_file.close()
# for a in range (len(sheet_lists)):
#     df = pd.read_excel(excel_file, sheet_name=sheet_lists[a], dtype=str)
#     header = df.columns.values #put tabel header to list
#     script = open(script_file_path, "a") #open/create script
#     script.write('create table '+sheet_lists[a]+'\n(\n')
#     for i in range (len(header)): #for loop that write header row to script
#         script.write('\t'+header[i])
#         for k in range (len(df)): #for loop that write one row of data to script
#             col_len_max = 0
#             for j in range (df.shape[1]): #df.shape = row length, itterate till last column in a row
#                 colmn=str(df.iloc[k,j]) #iloc locate data on k row, j column
#                 print(len(colmn))
#                 if col_len_max <= len(colmn):
#                     col_len_max = len(colmn)
#                 script.write(col_len_max)
#         if i == (len(header)-1):
#             script.write('\n)\n\n')
#         else:
#             script.write('\n')

# TODO: add create table to the script
# for a in range(len(sheet_lists)):
#     df = pd.read_excel(excel_file, sheet_name=sheet_lists[a], dtype=str)
#     # print(df)
#     header = df.columns.values  # put tabel header to list
#     script = open(script_file_path, "a")  # open/create script
#     script.write("create table " + sheet_lists[a] + " (\n")
#     for i in range(len(header)):  # for loop that write header row to script
#         script.write(header[i])
#         if i == (len(header) - 1):
#             script.write(" TIPE_DATA(ANGKA)\n);\n\n")
#         else:
#             script.write(" TIPE_DATA(ANGKA),\n")

for a in range(len(sheet_lists)):
    df = pd.read_excel(excel_file, sheet_name=sheet_lists[a], dtype=str)
    # print(df)
    header = df.columns.values  # put tabel header to list
    script = open(
        script_file_path, "w"
    )  # open/create script, file is overwitten, use a for append
    script.write("insert into " + sheet_lists[a] + " (")
    for i in range(len(header)):  # for loop that write header row to script
        script.write(header[i])
        if i == (len(header) - 1):
            script.write(") values\n")
        else:
            script.write(", ")

    for i in range(len(df)):  # for loop that write one row of data to script
        script.write("(")
        for j in range(
            df.shape[1]
        ):  # df.shape = row length, itterate till last column in a row
            colmn = str(df.iloc[i, j])  # iloc locate data on i row, j column
            if (
                colmn == "nan" or colmn == "NaT"
            ):  # empty data check, if empty write null instead nan/nat
                script.write("'null'")
            else:
                script.write("'" + colmn + "'")  # main write script
            if (
                j == df.shape[1] - 1
            ):  # finising check, if at the end of data write ')' instead ', )'
                break
            else:
                script.write(", ")
        if i == len(df) - 1:  # finishing check
            script.write(");\n")
        else:
            script.write("),")
        script.write("\n")
    script.close()
    print(str(i + 1) + " row created")

for a in range(len(sheet_lists)):
    df = pd.read_excel(excel_file, sheet_name=sheet_lists[a], dtype=str)
    # print(df)
    header = df.columns.values  # put tabel header to list
    script = open(script_file_path, "a")  # open/create script
    script.write("select * from " + sheet_lists[a] + "\n")
