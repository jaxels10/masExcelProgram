import pandas as pd
from Product import Product
from collections import defaultdict
import xlsxwriter
from excel_product import excel_product
from tkinter import *
from tkinter import filedialog


def read_file(file):
    return pd.read_excel(file)


def get_sorted_products():
    product_array = []
    current_product = None

    for i in range(len(file)):
        current_product = Product(file.loc[i][0], file.loc[i][1], file.loc[i][2], file.loc[i][3], file.loc[i][4],
                                 file.loc[i][5], file.loc[i][6], file.loc[i][7], file.loc[i][8], file.loc[i][9],
                                 file.loc[i][10], file.loc[i][11], file.loc[i][12], file.loc[i][13], file.loc[i][14],
                                 file.loc[i][15], file.loc[i][16], file.loc[i][17])
        product_array.append(current_product)

    groups = defaultdict(list)

    for obj in product_array:
        groups[obj.item_number].append(obj)

    return groups.values()



window = Tk()
file = filedialog.askopenfilename()
print(file)

T = Text(window, height=2, width=30)
T.pack()
T.insert(END, "Du kan nu lukke vinduet og vente på filen kommer frem")
window.title("Welcome to Jesper app")
window.geometry("600x200")



window.mainloop()


read_file = read_file(file)
column_names = read_file.columns.values
read_file.to_csv (r'DeleteThisFile.csv', index = None, header=True)
file = pd.read_csv('DeleteThisFile.csv')
sorted_product_array = list(get_sorted_products())

excel_product_array = []
row_index = 2
for product in sorted_product_array:
    excel_product_array.append(excel_product(row_index, product))
    row_index = row_index + 7












workbook = xlsxwriter.Workbook('rapport.xlsx', {'nan_inf_to_errors': True})
worksheet = workbook.add_worksheet("rapport")
top_line = workbook.add_format()
red_color = workbook.add_format({'font_color': '#FF0000'})
grey_color = workbook.add_format({'font_color': '#DCDCDC'})
top_line.set_bottom(5)
worksheet.freeze_panes(2, 0)
divider_line = workbook.add_format()
divider_line.set_bottom(1)
align_center = workbook.add_format()
align_center.set_align('center')

worksheet.write('A1', 'Quantity')
worksheet.write('E1', 'Month')
worksheet.write('A2', column_names[2], top_line)
worksheet.write('B2', column_names[0], top_line)
worksheet.write('C2', 'Name', top_line)
worksheet.write('D2', 'Type', top_line)
worksheet.write('E2', column_names[5], top_line)
worksheet.write('F2', column_names[6], top_line)
worksheet.write('G2', column_names[7], top_line)
worksheet.write('H2', column_names[8], top_line)
worksheet.write('I2', column_names[9], top_line)
worksheet.write('J2', column_names[10], top_line)
worksheet.write('K2', column_names[11], top_line)
worksheet.write('L2', column_names[12], top_line)
worksheet.write('M2', column_names[13], top_line)
worksheet.write('N2', column_names[14], top_line)
worksheet.write('O2', column_names[15], top_line)
worksheet.write('P2', column_names[16], top_line)
worksheet.write('Q2', column_names[17], top_line)
worksheet.write('R2', 'PURCHASES TO FC', top_line)

#Skriver type
for i in range(3, len(excel_product_array) * 6, 6):
    worksheet.write('D' + str(i), 'Forecast')
    worksheet.write('D' + str(i + 1), 'Actual Inv')
    worksheet.write('D' + str(i + 2), 'Sales Order')
    worksheet.write('D' + str(i + 3), 'PO')
    worksheet.write('D' + str(i + 4), 'Stock')
    worksheet.write('D' + str(i + 5), 'FC CARRYOVER STOCK', divider_line)

index = 3

excel_product_array.sort(key=lambda x: x.brand, reverse=False)


for product in excel_product_array:
    if product.brand is not None:

        worksheet.write('A' + str(index), product.brand)

    if product.code is not None:
        worksheet.write('B' + str(index), product.code)

    if product.description is not None:
        worksheet.write('C' + str(index), product.description)

    if product.ptf is not None:
        worksheet.write('R' + str(index), str(int(product.ptf)) + '%', align_center)

    if product.forecast_product is not None:
        worksheet.write('E' + str(index), product.forecast_product.month_1)
        worksheet.write('F' + str(index), product.forecast_product.month_2)
        worksheet.write('G' + str(index), product.forecast_product.month_3)
        worksheet.write('H' + str(index), product.forecast_product.month_4)
        worksheet.write('I' + str(index), product.forecast_product.month_5)
        worksheet.write('J' + str(index), product.forecast_product.month_6)
        worksheet.write('K' + str(index), product.forecast_product.month_7)
        worksheet.write('L' + str(index), product.forecast_product.month_8)
        worksheet.write('M' + str(index), product.forecast_product.month_9)
        worksheet.write('N' + str(index), product.forecast_product.month_10)
        worksheet.write('O' + str(index), product.forecast_product.month_11)
        worksheet.write('P' + str(index), product.forecast_product.month_12)
        worksheet.write('Q' + str(index), product.total_forecast)

    if product.invoiced_product is not None:
        worksheet.write('E' + str(index + 1), product.invoiced_product.month_1)
        worksheet.write('F' + str(index + 1), product.invoiced_product.month_2)
        worksheet.write('G' + str(index + 1), product.invoiced_product.month_3)
        worksheet.write('H' + str(index + 1), product.invoiced_product.month_4)
        worksheet.write('I' + str(index + 1), product.invoiced_product.month_5)
        worksheet.write('J' + str(index + 1), product.invoiced_product.month_6)
        worksheet.write('K' + str(index + 1), product.invoiced_product.month_7)
        worksheet.write('L' + str(index + 1), product.invoiced_product.month_8)
        worksheet.write('M' + str(index + 1), product.invoiced_product.month_9)
        worksheet.write('N' + str(index + 1), product.invoiced_product.month_10)
        worksheet.write('O' + str(index + 1), product.invoiced_product.month_11)
        worksheet.write('P' + str(index + 1), product.invoiced_product.month_12)
        worksheet.write('Q' + str(index + 1), product.total_invoiced)

    if product.sales_order_product is not None:
        worksheet.write('E' + str(index + 2), product.sales_order_product.month_1)
        worksheet.write('F' + str(index + 2), product.sales_order_product.month_2)
        worksheet.write('G' + str(index + 2), product.sales_order_product.month_3)
        worksheet.write('H' + str(index + 2), product.sales_order_product.month_4)
        worksheet.write('I' + str(index + 2), product.sales_order_product.month_5)
        worksheet.write('J' + str(index + 2), product.sales_order_product.month_6)
        worksheet.write('K' + str(index + 2), product.sales_order_product.month_7)
        worksheet.write('L' + str(index + 2), product.sales_order_product.month_8)
        worksheet.write('M' + str(index + 2), product.sales_order_product.month_9)
        worksheet.write('N' + str(index + 2), product.sales_order_product.month_10)
        worksheet.write('O' + str(index + 2), product.sales_order_product.month_11)
        worksheet.write('P' + str(index + 2), product.sales_order_product.month_12)
        worksheet.write('Q' + str(index + 2), product.total_sales_order)

    if product.buy_order_product is not None:
        worksheet.write('E' + str(index + 3), product.buy_order_product.month_1)
        worksheet.write('F' + str(index + 3), product.buy_order_product.month_2)
        worksheet.write('G' + str(index + 3), product.buy_order_product.month_3)
        worksheet.write('H' + str(index + 3), product.buy_order_product.month_4)
        worksheet.write('I' + str(index + 3), product.buy_order_product.month_5)
        worksheet.write('J' + str(index + 3), product.buy_order_product.month_6)
        worksheet.write('K' + str(index + 3), product.buy_order_product.month_7)
        worksheet.write('L' + str(index + 3), product.buy_order_product.month_8)
        worksheet.write('M' + str(index + 3), product.buy_order_product.month_9)
        worksheet.write('N' + str(index + 3), product.buy_order_product.month_10)
        worksheet.write('O' + str(index + 3), product.buy_order_product.month_11)
        worksheet.write('P' + str(index + 3), product.buy_order_product.month_12)
        worksheet.write('Q' + str(index + 3), product.total_buy_order)

    if product.stock_product is not None:
        worksheet.write('E' + str(index + 4), product.stock_product.month_1)
        worksheet.write('F' + str(index + 4), product.stock_product.month_2)
        worksheet.write('G' + str(index + 4), product.stock_product.month_3)
        worksheet.write('H' + str(index + 4), product.stock_product.month_4)
        worksheet.write('I' + str(index + 4), product.stock_product.month_5)
        worksheet.write('J' + str(index + 4), product.stock_product.month_6)
        worksheet.write('K' + str(index + 4), product.stock_product.month_7)
        worksheet.write('L' + str(index + 4), product.stock_product.month_8)
        worksheet.write('M' + str(index + 4), product.stock_product.month_9)
        worksheet.write('N' + str(index + 4), product.stock_product.month_10)
        worksheet.write('O' + str(index + 4), product.stock_product.month_11)
        worksheet.write('P' + str(index + 4), product.stock_product.month_12)

    if product.fcs is not None:
        worksheet.write('A' + str(index + 5), '', divider_line)
        worksheet.write('B' + str(index + 5), '', divider_line)
        worksheet.write('C' + str(index + 5), '', divider_line)
        worksheet.write('E' + str(index + 5), product.fcs.month_1, divider_line)
        worksheet.write('F' + str(index + 5), product.fcs.month_2, divider_line)
        worksheet.write('G' + str(index + 5), product.fcs.month_3, divider_line)
        worksheet.write('H' + str(index + 5), product.fcs.month_4, divider_line)
        worksheet.write('I' + str(index + 5), product.fcs.month_5, divider_line)
        worksheet.write('J' + str(index + 5), product.fcs.month_6, divider_line)
        worksheet.write('K' + str(index + 5), product.fcs.month_7, divider_line)
        worksheet.write('L' + str(index + 5), product.fcs.month_8, divider_line)
        worksheet.write('M' + str(index + 5), product.fcs.month_9, divider_line)
        worksheet.write('N' + str(index + 5), product.fcs.month_10, divider_line)
        worksheet.write('O' + str(index + 5), product.fcs.month_11, divider_line)
        worksheet.write('P' + str(index + 5), product.fcs.month_12, divider_line)
        worksheet.write('Q' + str(index + 5), product.fcs.month_12, divider_line)
        worksheet.write('R' + str(index + 5), '', divider_line)


    index = index + 6








worksheet.conditional_format(
    'A1:U10000', {
        'type': 'cell',
        'criteria': '<',
        'value': 0,
        'format': red_color
    })
worksheet.conditional_format(
    'A1:U10000', {
        'type': 'cell',
        'criteria': '=',
        'value': 0,
        'format': grey_color
    })



worksheet.set_column(0, 0, 18)
worksheet.set_column(1, 1, 17)
worksheet.set_column(2, 2, 50)
worksheet.set_column(3, 3, 20)
worksheet.set_column(4, 16, 9)
worksheet.set_column(17, 17, 20)
workbook.close()

# new_list[0][0].name


