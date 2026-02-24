import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import plotly.graph_objects as go
import os
import io

os.getcwd()

df = pd.read_csv(r'C:\Users\kevin.yu\Documents\Documents\git_repo_1\py_demo\superstore_demo\Sample_Superstore.csv', encoding='latin1')

df.drop_duplicates(inplace=True)

preview = df.head()
dataset = df.info()
summary = df.describe()
checknull = df.isnull().sum()

df['Profit Margin'] = np.where(df['Sales'] != 0, df['Profit'] / df['Sales'], 0)
df.rename(columns={'Discount': 'Discount Rate'}, inplace=True)
df['COGS'] = np.where(df['Sales'] != 0, df['Sales'] - df['Profit'], 0)

region_summary = df.groupby('Region')[['Sales', 'Profit']].sum().reset_index()
region_summary['Profit Margin'] = np.where(region_summary['Sales'] != 0, region_summary['Profit'] / region_summary['Sales'], 0)

top_profit = df.sort_values(by='Profit', ascending=False)
top_profit = top_profit[['Order ID', 'Product Name', 'Profit']].head(10)

top_sales = df.groupby('Product Name')['Sales'].sum().sort_values(ascending=False).head(10).reset_index()

plt.figure(figsize=(10, 6))
plt.barh(top_sales['Product Name'], top_sales['Sales'])
plt.title('Top 10 Most Sales Products')
plt.xlabel('Sales')
plt.gca().invert_yaxis()
plt.tight_layout()
bar_buffer = io.BytesIO()
plt.savefig(bar_buffer, format='png')
bar_buffer.seek(0)


sankey_df = df.groupby(['Category', 'Sub-Category']).agg(
    Sales=('Sales', 'sum'),
    COGS=('COGS', 'sum'),
    Profit=('Profit', 'sum')
).reset_index()

categories = sankey_df['Category'].unique().tolist()
sub_categories = sankey_df['Sub-Category'].unique().tolist()
right_nodes = ['Sales', 'Profit', 'COGS']

all_nodes = [str(x) for x in sub_categories + categories + right_nodes]

def idx(name): 
    return all_nodes.index(name)

sources, targets, values, colors = [], [], [], []

# Sub-Cat to Cat
for _, row in sankey_df.iterrows():
    sources.append(idx(row['Sub-Category']))
    targets.append(idx(row['Category']))
    values.append(row['Sales'])
    colors.append('rgba(70, 130, 180, 0.4)')

#Cat to P&L
cat_sums = sankey_df.groupby('Category').agg(
    Sales=('Sales', 'sum'),
    COGS=('COGS', 'sum'),
    Profit=('Profit', 'sum')
).reset_index()

for _, row in cat_sums.iterrows():
    # Category → Profit
    sources.append(idx(row['Category']))
    targets.append(idx('Profit'))
    values.append(row['Profit'])
    colors.append('rgba(60, 179, 113, 0.4)')

    # Category → Discount Amount
    sources.append(idx(row['Category']))
    targets.append(idx('COGS'))
    values.append(row['COGS'])
    colors.append('rgba(220, 20, 60, 0.4)')

node_colors = (
    ['rgba(70, 130, 180, 0.8)'] * len(sub_categories) +   # steelblue - matches Sub-Cat → Category links
    ['rgba(70, 130, 180, 0.8)'] * len(categories) +        # steelblue - matches Category node
    ['rgba(70, 130, 180, 0.8)'] +                           # Sales
    ['rgba(60, 179, 113, 0.8)'] +                           # Profit - matches green links
    ['rgba(220, 20, 60, 0.8)']                              # COGS - matches red links
)

# Pre-calculate node values
sub_cat_sales = sankey_df.groupby('Sub-Category')['Sales'].sum()
cat_sales = sankey_df.groupby('Category')['Sales'].sum()
right_values = [sankey_df['Sales'].sum(), sankey_df['Profit'].sum(), sankey_df['COGS'].sum()]

# Build labeled node names
labeled_nodes = (
    [f'{n} ${sub_cat_sales[n]/1000:,.0f}K' for n in sub_categories] +
    [f'{n} ${cat_sales[n]/1000:,.0f}K' for n in categories] +
    [f'Sales<br>${right_values[0]/1000:,.0f}K',
     f'Profit<br>${right_values[1]/1000:,.0f}K',
     f'COGS<br>${right_values[2]/1000:,.0f}K']
)

# Build Sankey Visual
fig_sankey = go.Figure(go.Sankey(
    arrangement='fixed',
    node=dict(
        pad=10,
        thickness=30,
        line=dict(color='black', width=0.5),
        label=labeled_nodes,
        color=node_colors,
        x=[*[0.01]*len(sub_categories), *[0.5]*len(categories), 0.99, 0.99, 0.99],
        y=[*[None]*len(sub_categories), *[None]*len(categories), 0.10, 0.5, 0.99]
    ),
    link=dict(
        source=sources,
        target=targets,
        value=values,
        color=colors,
        label=[f'${v:,.0f}' for v in values]         # Visible label on each link
    )
))

fig_sankey.update_layout(
    title_text='Superstore Product P&L',
    font_size=11,
    width=1400,
    height=700
)

# Buffer for excel
sankey_buffer = io.BytesIO()
fig_sankey.write_image(sankey_buffer, format='png', scale=2)
sankey_buffer.seek(0)

plt.savefig('bar_chart.png', dpi=150)
fig_sankey.write_html('sankey_chart.html')
fig_sankey.write_image('sankey_chart.png', scale=2)
#fig_sankey.show()


with pd.ExcelWriter(
    'Superstore Analysis Output.xlsx', #Path.
    engine='xlsxwriter', #Engine to use for writing. If None, defaults to io.excel.<extension>.writer. It can only be passed as a keyword argument.
    date_format='YYYY-MM-DD', #default none
    datetime_format='YYYY-MM-DD HH:MM:SS', #default none
    mode='w' #mode{‘w’, ‘a’}, default ‘w’ write or append
    #storage_options=None,
    #if_sheet_exists='new' #{‘error’, ‘new’, ‘replace’, ‘overlay’}, default ‘error’
    ) as writer:

    summary.to_excel(
        excel_writer=writer,
        sheet_name='overview',
        na_rep='NaN',
        #index=True,
        index_label='Index',
        startrow=1,
        startcol=1
    ) 

    workbook = writer.book
    chart_1 = workbook.add_worksheet('Chart_1')
    chart_1.insert_image('B2', 'chart.png', {'image_data': bar_buffer})
    chart_2 = workbook.add_worksheet('Chart_2')
    chart_2.insert_image('B2', 'sankey.png', {'image_data': sankey_buffer})

    sankey_df.to_excel(
        excel_writer=writer, #File path or existing ExcelWriter.
        sheet_name='sankey_df', #Name of sheet which will contain DataFrame.
        na_rep='', #null values representation
        float_format='%.2f', #decimal formatting
        index=True, #add index 
        index_label='Index', #index column name
        startrow=0, #starting row value
        startcol=0 #starting column value
    )

    df.to_excel(
        excel_writer=writer, #File path or existing ExcelWriter.
        sheet_name='raw', #Name of sheet which will contain DataFrame.
        na_rep='', #null values representation
        float_format='%.2f', #decimal formatting
        index=True, #add index 
        index_label='Index', #index column name
        startrow=0, #starting row value
        startcol=0 #starting column value
    )