import xlwings as xw

print('Finish importing package')

#Open file and active sheet 1
wb_main = xw.Book(r'D:\python\Saoke\#main.xlsx').sheets[0]
wb_bs = xw.Book(r'D:\python\Saoke\#main_bs.xlsx').sheets[0]


# wb_main = xw.Book(r'D:\python\Saoke\#346375_1.xlsx').sheets[0]
# wb_bs = xw.Book(r'D:\python\Saoke\#346375_1_bs.xlsx').sheets[0]

#Count row for each file
rows_main = wb_main.range('A1').current_region.last_cell.row
rows_bs = wb_bs.range('A1').current_region.last_cell.row

list_tracsaction_main = wb_main.range(f'A2:A{rows_main+1}').value

print(rows_main)
print(rows_bs)

rows_append = rows_main
list_columns = ['D', 'E', 'F', 'G', 'H', 'I']

for row_bs in range(2,rows_bs+1):
    transaction_bs = wb_bs.range(f'A{row_bs}').value
    if transaction_bs not in list_tracsaction_main:
        rows_append +=1
        wb_main.range(f'A{rows_append}').value = wb_bs.range(f'A{row_bs}').value
        for i in list_columns:
            wb_main.range(f'{i}{rows_append}').value = wb_bs.range(f'{i}{row_bs}').value
    else:
        for row_main in range(2,rows_main+1):
            transaction_main = wb_main.range(f'A{row_main}').value

            if transaction_bs == transaction_main:
                print(f'A{row_bs} = A{row_main}')

                for col in list_columns:
                    if wb_bs.range(f'{col}{row_bs}').value is None or wb_main.range(f'{col}{row_main}').value == wb_bs.range(f'{col}{row_bs}').value:
                        continue
                    elif wb_main.range(f'{col}{row_main}').value is None:
                        wb_main.range(f'{col}{row_main}').value = wb_bs.range(f'{col}{row_bs}').value
                        wb_main[f'{col}{row_main}'].color = "#ff0000"
                    else:
                        rows_append +=1
                        wb_main.range(f'A{rows_append}').value = wb_bs.range(f'A{row_bs}').value
                        for i in list_columns:
                            wb_main.range(f'{i}{rows_append}').value = wb_bs.range(f'{i}{row_bs}').value
                        wb_main[f'A{rows_append}:I{rows_append}'].color = "#ff0000"   


            


            




            

        
    
# print(row)
# print(row_bs)

# row_main = 2
# while wb_main.range(f'A{row_main}').value:
#     transaction_main = wb_main.range(f'A{row_main}').value
#     row_bs = 2
#     while wb_bs.range(f'A{row_bs}').value:
#         transaction_bs = wb_bs.range(f'A{row_bs}').value
        
#     row_main +=1
        

# Change color cell
# wb_main['A10'].color = "#ff0000"
    
# Change color text
# wb_main['A10'].font.color = "#ff0000"

# for i in range(3):
#     if i == 0:
#         continue
#     elif i == 1:
#         print(i)
#     else:
#         print(i)





