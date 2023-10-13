# The given list of strings
data_list = ['Design Name', 'Design Revision', 'Design PartNumber', 'Wire Name', 'Marker spaced', 'Marker middle', 'Marker End1', 'Marker End2', 'INT part number', 'Multicore name', 'Color', 'Type', 'CSA', 'Length', 'Option', 'From Conn1', 'From Pin1', 'To Conn2', 'To Pin2', 'INT Term1', 'Strip length1', 'INT Term2', 'Strip length2', 'CUST Term1', 'CUST Term2', 'SUPP Term1', 'SUPP Term2', 'INT Seal1', 'INT Seal2', 'CUST Seal1', 'CUST Seal2', 'SUPP Seal1', 'SUPP Seal2', 'FM Code']

# String to check
string_to_check = 'Wire Name'

# Check if the string is in the list
if string_to_check in data_list:
    print(f"'{string_to_check}' is in the list.")
else:
    print(f"'{string_to_check}' is not in the list.")
