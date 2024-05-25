import os
import openpyxl
import msoffcrypto
import io

# Get the current directory
current_dir = os.getcwd()
sum_value_total = 0.0

# Iterate over all files in the directory
for filename in os.listdir(current_dir):
    # Check if the file is an xlsx file
    if filename.endswith(".xlsx"):
        # Construct the full path to the file
        file_path = os.path.join(current_dir, filename)

        decrypted_workbook = io.BytesIO()

        with open(file_path, 'rb') as file:
            office_file = msoffcrypto.OfficeFile(file)
            office_file.load_key(password='667287')
            office_file.decrypt(decrypted_workbook)
            print(file.name)

        # Open the workbook using openpyxl
        workbook = openpyxl.load_workbook(filename=decrypted_workbook, data_only=True)
        
        # Get the first sheet of the workbook
        sheet = workbook.active

        # Get the value of the cell I1
        cell_value = sheet['I1'].value

        # Keep increasing the row index until the value is not "ยอดรวมทั้งหมด"
        row_index = 1
        while cell_value != "ยอดรวมทั้งหมด":
            row_index += 1
            cell_value = sheet[f'I{row_index}'].value

        # Sum the values from J18 to J[row_index]
        sum_value = 0.0
        for i in range(18, row_index + 1):
            j_value = sheet[f'J{i}'].value
            if j_value is not None:
                sum_value += float(j_value)

        # Print the sum value to the terminal
        print(sum_value)
        sum_value_total += sum_value

# Print the total sum value to the terminal
print('----')
print(sum_value_total)

