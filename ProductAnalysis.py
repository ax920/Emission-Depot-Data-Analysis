import pandas as pd
import tkinter
import tkinter.filedialog
import sys

# OPEN FILE EXPLORER

# root = tkinter.Tk()
# root.withdraw()  # Get rid of extra tkinter window

# filename = tkinter.filedialog.askopenfilename()

# print(filename)
# if filename == '':
#     sys.exit()

'''
EXAMPLE TO GET SPECIFIC CELL
 print(pd.isnull(sales.iloc[1][0]))
 print(sales.iloc[1][0])
'''

# START SCRIPT
# sales = pd.read_excel(filename)
sales = pd.read_excel('sales_by_product.xlsx')
sales = sales.drop(sales.index[0:5])
sales = sales.reset_index(drop=True)  # reset index since we drop rows
sales = sales.rename(
    columns={'Emissions Depot': 'Item Name',
             'Unnamed: 1': 'Date',
             'Unnamed: 2': 'Transaction Type',
             'Unnamed: 3': 'Transction Number',
             'Unnamed: 4': 'Customer Name',
             'Unnamed: 5': 'Item Description',
             'Unnamed: 6': 'Quantity',
             'Unnamed: 7': 'Sales Price',
             'Unnamed: 8': 'Revenue',
             'Unnamed: 9': 'Balance'})

# Remove all rows that start with 'Total'
sales = sales[~sales["Item Name"].str.contains("Total for", na=False)]
sales = sales.reset_index(drop=True)  # reset index since we drop rows

# Remove all shipping entries

shippingIndex = sales.loc[sales['Item Description'] == 'Shipping'].index
shippingIndex = shippingIndex[0] - 1
sales = sales[:shippingIndex]

# Replace all NaN Item Names with appropriate Item Name
# by looping through all entries

current_item_name = ''
numRows = sales.shape[0]

for x in range(0, numRows):
    i = sales.iloc[x][0]
    isNull = pd.isnull(i)
    if not(isNull):
        current_item_name = i
    else:
        sales.at[x, 'Item Name'] = current_item_name

# Drop all dupicate Item Name Header rows

sales = sales[pd.notnull(sales['Date'])]
sales = sales.reset_index(drop=True)  # reset index since we drop rows

# Load in product list CSV

product_list = pd.read_excel('product_list.xls')
for x in range(0, product_list.shape[0]):
    # index of first ':', we want to remove the item category
    index = product_list.iloc[x]['Product/Service Name'].find(':') + 1
    if index != 0:
        name = product_list.iloc[x]['Product/Service Name']
        product_list.at[x, 'Product/Service Name'] = name[index:]
        # print(product_list.iloc[x]['Product/Service Name'])

# Item Name - Sell Price dictionary
# Item Name - Purchase Price dictionary
sell_price_dict = {}
purchase_price_dict = {}

for x in range(0, product_list.shape[0]):
    sell_price_dict[product_list.iloc[x]['Product/Service Name']
                    ] = product_list.iloc[x]['Sales Price / Rate']

for x in range(0, product_list.shape[0]):
    purchase_price_dict[product_list.iloc[x]['Product/Service Name']
                        ] = product_list.iloc[x]['Purchase Cost']

# Keep only these columns from sales DataFrame
summary = sales.loc[:, ['Item Name',
                        'Quantity',
                        'Revenue']]
summary = summary.groupby(
    ['Item Name']).sum().reset_index()
summary = summary.reset_index(drop=True)  # reset index since we drop cols

# summary['Purchase Price'] = 0
summary['Item Name'] = summary['Item Name'].str.lstrip()

# summary['Item Name'].apply(lambda item: purchase_price_dict(item))
summary['Purchase Price'] = summary['Item Name'].map(
    purchase_price_dict, na_action='ignore')

summary['Expenses'] = summary['Quantity'] * summary['Purchase Price']
summary['Net Profit'] = summary['Revenue'] - summary['Expenses']
summary['Profit Percentage'] = (summary['Revenue'] -
                                summary['Expenses']) / summary['Revenue'] * 100

summary = summary.round(decimals=2)
summary = summary.sort_values(by=['Net Profit'], ascending=False)

print("summary", summary)

# END SCRIPT

summary.to_excel('sales_output.xlsx')
