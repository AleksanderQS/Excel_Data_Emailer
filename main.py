import pandas as pd
import win32com.client as win32

sales_sheet = pd.read_excel('Sales.xlsx')
pd.set_option('display.max_columns', None)

profit = sales_sheet[['ID Store', 'Final price']].groupby('ID Store').sum()
quantity = sales_sheet[['ID Store', 'Quantity']].groupby('ID Store').sum()
average = (profit['Final price'] / quantity['Quantity']).to_frame()
average = average.rename(columns={0: 'Average price'})

# Send email

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.to = 'your_email'
mail.Subject = 'Sales report'
mail.HTMLBody = f'''
<p>To whom it may concern,</p>

<p>Below is a sales report for each store.</p>

<p>Profit:</p>
{profit.to_html(formatters={'Final price': '${:,.2f}'.format})}

<p>Quantity:</p>
{quantity.to_html(formatters={'Quantity': '{:.0f}'.format})}

<p>Average price:</p>
{average.to_html(formatters={'Average price': '${:,.2f}'.format})}

<p>Kind regards,</p>

<p>Aleksander</p>
'''

mail.Send()
