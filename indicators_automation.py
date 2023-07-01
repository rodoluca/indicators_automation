#!/usr/bin/env python
# coding: utf-8

# ### Step 1 - Import Files and Libraries

# In[7]:


#import libs
import pandas as pd
import win32com.client as win32
import pathlib


# In[8]:


#import datasets
emails = pd.read_excel(r'stores_datasets\Emails.xlsx')
stores = pd.read_csv(r'stores_datasets\Stores.csv', encoding='latin1', sep=';')
sales = pd.read_excel(r'stores_datasets\Sales.xlsx')
display(emails)
display(stores)
display(sales)


# ### Step 2 - Create a Table for Each Store and Define the Indicator Day

# In[9]:


#Include store name in sales
sales = sales.merge(stores, on='Store ID')
display(sales)


# In[10]:


store_dictionary = {}
for store in stores['Store']:
    store_dictionary[store] = sales.loc[sales['Store']==store, :]
    
#test
display(store_dictionary['Rio Mar Recife'])
display(store_dictionary['Shopping Vila Velha'])


# In[11]:


indicator_date = sales['Date'].max()
print(indicator_date)
print('{}/{}'.format(indicator_date.day, indicator_date.month))


# In[ ]:





# ### Step 3 - Save the spreadsheet in the backup folder

# In[12]:


#Check if the folder already exists
backup_path = pathlib.Path(r'backup_stores_files')

backup_folder_files = backup_path.iterdir()
backup_folder_names = [file.name for file in backup_folder_files]

for store in store_dictionary:
    if store not in backup_folder_names:
        new_folder = backup_path / store
        new_folder.mkdir()
    
    # Save inside the folder
    file_name = '{}_{}_{}.xlsx'.format(indicator_date.month, indicator_date.day, store)
    file_path = backup_path / store / file_name
    store_dictionary[store].to_excel(file_path)


# ### Step 4 - Calculate the indicator for 1 store

# In[13]:


#Goal Definitions
daily_revenue_target = 1000
annual_revenue_target = 1650000
daily_product_count_target = 4
annual_product_count_target = 120
daily_average_ticket_target = 500
annual_average_ticket_target = 500


# ### Step 5 - Send email to the manager
# ### Step 6 - Automate all stores

# In[14]:


for store in store_dictionary:
    store_sales = store_dictionary[store]
    store_sales_day = store_sales.loc[store_sales['Date'] == indicator_date, :]

    # Revenue
    revenue_year = store_sales['Final Value'].sum()
    #print(revenue_year)
    revenue_day = store_sales_day['Final Value'].sum()
    #print(revenue_day)

    # Product diversity
    product_count_year = len(store_sales['Product'].unique())
    #print(product_count_year)
    product_count_day = len(store_sales_day['Product'].unique())
    #print(product_count_day)

    # Average ticket
    sale_value = store_sales.groupby('Sale Code').sum()
    average_ticket_year = sale_value['Final Value'].mean()
    #print(average_ticket_year)
    # Average ticket (daily)
    sale_value_day = store_sales_day.groupby('Sale Code').sum()
    average_ticket_day = sale_value_day['Final Value'].mean()
    #print(average_ticket_day)

    # Send email
    outlook = win32.Dispatch('outlook.application')

    name = emails.loc[emails['Store'] == store, 'Manager'].values[0]
    mail = outlook.CreateItem(0)
    mail.To = emails.loc[emails['Store'] == store, 'Email'].values[0]
    mail.Subject = f'OnePage {indicator_date.day}/{indicator_date.month} - Store {store}'
    #mail.Body = 'Email Body'

    if revenue_day >= daily_revenue_target:
        revenue_day_color = 'green'
    else:
        revenue_day_color = 'red'
    if revenue_year >= annual_revenue_target:
        revenue_year_color = 'green'
    else:
        revenue_year_color = 'red'
    if product_count_day >= daily_product_count_target:
        product_count_day_color = 'green'
    else:
        product_count_day_color = 'red'
    if product_count_year >= annual_product_count_target:
        product_count_year_color = 'green'
    else:
        product_count_year_color = 'red'
    if average_ticket_day >= daily_average_ticket_target:
        average_ticket_day_color = 'green'
    else:
        average_ticket_day_color = 'red'
    if average_ticket_year >= annual_average_ticket_target:
        average_ticket_year_color = 'green'
    else:
        average_ticket_year_color = 'red'

    mail.HTMLBody = f'''
    <p>Good morning, {name}</p>

    <p>The result for yesterday <strong>({indicator_date.day}/{indicator_date.month})</strong> for <strong>Store {store}</strong> was:</p>

    <table>
      <tr>
        <th>Indicator</th>
        <th>Value (Day)</th>
        <th>Target (Day)</th>
        <th>Scenario (Day)</th>
      </tr>
      <tr>
        <td>Revenue</td>
        <td style="text-align: center">R${revenue_day:.2f}</td>
        <td style="text-align: center">R${daily_revenue_target:.2f}</td>
        <td style="text-align: center"><font color="{revenue_day_color}">◙</font></td>
      </tr>
      <tr>
        <td>Product Diversity</td>
        <td style="text-align: center">{product_count_day}</td>
        <td style="text-align: center">{daily_product_count_target}</td>
        <td style="text-align: center"><font color="{product_count_day_color}">◙</font></td>
      </tr>
      <tr>
        <td>Average Ticket</td>
    <td style="text-align: center">$ {average_ticket_year:.2f}</td>
    <td style="text-align: center">$ {annual_average_ticket_target:.2f}</td>
    <td style="text-align: center"><font color="{average_ticket_year_color}">◙</font></td>
  </tr>
</table>

<p>Please find attached the spreadsheet with all the data for more details.</p>

<p>If you have any questions, feel free to contact me.</p>
<p>Regards, Lucas</p>
'''

    # Attachments (you can add as many as you want):
    attachment = pathlib.Path.cwd() / backup_path / store / f'{indicator_date.month}_{indicator_date.day}_{store}.xlsx'
    mail.Attachments.Add(str(attachment))

    mail.Send()
    print('Email for Store {} sent'.format(store))


# In[ ]:





# 

# ### Step 7 - Create ranking for management

# In[15]:


revenue_stores = sales.groupby('Store')[['Store', 'Final Value']].sum()
revenue_stores_year = revenue_stores.sort_values(by='Final Value', ascending=False)
display(revenue_stores_year)

file_name = '{}_{}_Annual Ranking.xlsx'.format(indicator_date.month, indicator_date.day)
revenue_stores_year.to_excel(r'backup_stores_files\{}'.format(file_name))

sales_day = sales.loc[sales['Date'] == indicator_date, :]
revenue_stores_day = sales_day.groupby('Store')[['Store', 'Final Value']].sum()
revenue_stores_day = revenue_stores_day.sort_values(by='Final Value', ascending=False)
display(revenue_stores_day)

file_name = '{}_{}_Daily Ranking.xlsx'.format(indicator_date.month, indicator_date.day)
revenue_stores_day.to_excel(r'backup_stores_files\{}'.format(file_name))


# ### Step 8 - Send email to management

# In[16]:


#Send the email
outlook = win32.Dispatch('outlook.application')

mail = outlook.CreateItem(0)
mail.To = emails.loc[emails['Store'] == 'Diretoria', 'Email'].values[0]
mail.Subject = f'Daily Ranking {indicator_date.day}/{indicator_date.month}'
mail.Body = f'''
Dear all,

Best store of the Day in Revenue: Store {revenue_stores_day.index[0]} with Revenue $ {revenue_stores_day.iloc[0, 0]:.2f}
Worst store of the Day in Revenue: Store {revenue_stores_day.index[-1]} with Revenue $ {revenue_stores_day.iloc[-1, 0]:.2f}

Best store of the Year in Revenue: Store {revenue_stores_year.index[0]} with Revenue $ {revenue_stores_year.iloc[0, 0]:.2f}
Worst store of the Year in Revenue: Store {revenue_stores_year.index[-1]} with Revenue $ {revenue_stores_year.iloc[-1, 0]:.2f}

Attached are the rankings of all stores for the year and the day.

If you have any questions, feel free to reach out.

Best regards,

Lucas
'''

# Attachments
attachment = pathlib.Path.cwd() / backup_path / f'{indicator_date.month}_{indicator_date.day}_Annual Ranking.xlsx'
mail.Attachments.Add(str(attachment))
attachment = pathlib.Path.cwd() / backup_path / f'{indicator_date.month}_{indicator_date.day}_Daily Ranking.xlsx'
mail.Attachments.Add(str(attachment))


mail.Send()
print('Email sent to the Management')


# In[ ]:





# In[ ]:




