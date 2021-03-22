from openpyxl import load_workbook
import pandas as pd


class my_class:

    def Read_Excel_File(self):
        # to take the file name

        input_file = input("Enter file name: ")
        file_name = pd.ExcelFile(input_file)
        name_of_sheets = file_name.sheet_names

        num_of_sheets = len(name_of_sheets)

        # loop for printing the name of sheets in file

        print('\n')
        for i in range(num_of_sheets):
            print(i + 1, name_of_sheets[i])

        user_choice = int(input('\nIn which Sheet you want to search: '))

        sheet_df = pd.read_excel(file_name, sheet_name=name_of_sheets[user_choice - 1])
        head_of_sheet = sheet_df.head()
        heads_list = list(head_of_sheet)
        num_of_heads = len(heads_list)

        # loop for printing the name of rows in sheet

        print('\n')
        for j in range(num_of_heads):
            print(j + 1, heads_list[j])

        parameter = int(input("\nSelect the Parameter for search: "))

        row_df = pd.read_excel(file_name, sheet_name=name_of_sheets[user_choice - 1])
        row_data = row_df[[heads_list[parameter - 1]]]
        row_data_dtype = row_data.dtypes
        dtype_head = row_data_dtype.head()
        res5 = list(dtype_head)
        res6 = res5[0]

        res7 = (row_data.values.tolist())
        length_of_res7 = len(res7)

        my_list = []

        if res6 == 'int64':
            temp = int(input("\nEnter Data for you want to search: "))
            my_list.append(temp)

        elif res6 == 'object':
            temp = str(input("\nEnter Data for you want to search: "))
            my_list.append(temp)

        elif res6 == 'float64':
            temp = float(input("\nEnter Data for you want to search: "))
            my_list.append(temp)

        temp3 = res7.index(my_list)

        shiv = sheet_df.iloc[temp3]

        total_cols = len(shiv.axes[0])
        print(total_cols)

        path = file_name
        book = load_workbook(path)
        writer = pd.ExcelWriter(path, engine='openpyxl')
        writer.book = book
        if 'Mastersheet1' in book.sheetnames:
            pfd = book['Mastersheet1']
            book.remove(pfd)
        shiv.to_excel(writer, sheet_name='Mastersheet1')
        ws = book["Mastersheet1"]
        wcell1 = ws.cell(1,3)
        wcell1.value = total_cols

        writer.save()
        writer.close()


my_class.Read_Excel_File(self=0)
