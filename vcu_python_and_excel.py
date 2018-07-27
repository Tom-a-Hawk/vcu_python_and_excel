import pandas as pd
from pandas.api.types import CategoricalDtype
import os

list_of_excel_files = os.listdir()

# First we initialize an empty DataFrame.
# All of the data will be appended to df_agg_data
df_agg_data = pd.DataFrame()

# loop through our list of files
for file_name in list_of_excel_files:

    # check to see if file ends with 'xlsx'
    if file_name.endswith('xlsx') and 'sales' in file_name:
        # read in the excel file with pandas
        df = pd.read_excel(io=file_name, header=0)

        # append data to our empty DataFrame
        df_agg_data = df_agg_data.append(df)


months = ["january", "february", "march", "april", "may", "june", "july",
          "august", "september", "october", "november", "december"]

# let python know that months is categorical data, and that the order of months is important

cat_type = CategoricalDtype(categories=months, ordered=True)
df_agg_data['month'] = df_agg_data['month'].astype(cat_type)
df_agg_data = df_agg_data.sort_values(by='month')

df_agg_data_totals = df_agg_data.pivot_table(index=['product', 'month'], values='sales', aggfunc=sum)

df_agg_data_totals_unstacked = df_agg_data.pivot_table(index=['product', 'month'],
                                                       values='sales',
                                                       aggfunc=sum).unstack()

# january data - will use to make a bar chart
df_jan_totals = df_agg_data[df_agg_data['month']=='january']

products = list(df_agg_data['product'].unique())

writer = pd.ExcelWriter('python_output.xlsx', engine='xlsxwriter')

# write our unstacked pivot table to excel
df_agg_data_totals_unstacked.to_excel(writer, sheet_name='agg')  # no chart

for product in products:
    # subset the aggregated data by product
    df_product = df_agg_data[df_agg_data['product'] == product]
    # select the column we want to grab values from (sales, qty, etc.)
    df_product_subset = df_product[['month', 'sales']]
    # write the subset table to excel file
    df_product_subset.to_excel(writer, sheet_name=product, index=False)

    numer_of_rows_for_chart = len(df_product_subset)  # no hardcoding!

    workbook = writer.book
    worksheet = writer.sheets[product]  # naming each worksheet after the product name
    chart = workbook.add_chart({'type': 'line'})

    # https://xlsxwriter.readthedocs.io/index.html
    # [worksheet name, start row, start column, stop row, start column]
    chart.add_series({
        'categories': [product, 1, 0, numer_of_rows_for_chart, 0],
        'values': [product, 1, 1, numer_of_rows_for_chart, 1],
        'data_labels': {'value': True, 'position': 'below'},
        'marker': {'type': 'circle'}})

    chart.set_title({'name': product + " Sales"})

    chart.set_legend({'position': 'none'})

    chart.set_y_axis({'name': 'Sales in $'})

    chart.set_plotarea({
        'gradient': {'colors': ['#FFEFD1', '#F0EBD5', '#B69F66']}
    })

    worksheet.insert_chart('E2', chart)

# throwing in a bar chart for good measure
df_jan_totals = df_agg_data[df_agg_data['month'] == 'january']  # just january sales
df_jan_totals.to_excel(writer, sheet_name='jan_totals', index=False)  # write the dataframe to excel

# Create a new bar chart.
worksheet = writer.sheets['jan_totals']  # manually define the worksheet name this time
chart_bar = workbook.add_chart({'type': 'bar'})

chart_bar.add_series({
    'name': ['jan_totals', 1, 2],
    'categories': ['jan_totals', 1, 1, 4, 1],
    'values': ['jan_totals', 1, 2, 4, 2],
})

# Add a chart title and some axis labels.
chart_bar.set_title({'name': 'January - Quantity Sold'})
chart_bar.set_x_axis({'name': 'Number of Products Sold'})
chart_bar.set_y_axis({'name': 'Product Name'})
chart_bar.set_legend({'none': True})  # legen is useless in this scenario

# Set an Excel chart style.
chart_bar.set_style(11)

# Insert the chart into the worksheet (with an offset).
worksheet.insert_chart('E2', chart_bar, {'x_offset': 25, 'y_offset': 10})

# save the excel results
writer.save()