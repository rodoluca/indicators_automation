# indicators_automation

Objective: Create a comprehensive project involving the automation of a computer-based process.
Description:
Imagine that you work for a large clothing store network with 25 stores spread across Brazil.

Every morning, the data analysis team calculates the so-called One Pages and sends each store manager their respective One Page, along with all the information used to calculate the indicators.

A One Page is a very simple and straightforward summary used by the store management team to understand the key indicators of each store. It allows, on a single page (hence the name One Page), both the comparison between different stores and which indicators that store was able to achieve or not on that day.

OnePage example:
![image](https://github.com/rodoluca/indicators_automation/assets/115651087/0fe142de-6e55-48b4-aa2e-4e55eaf1fb46)

Your role as a Data Analyst is to create a highly automated process to calculate the One Page for each store and send an email to the respective store manager. The email should include the One Page in the body of the email and the complete data file for that specific store attached.

Example: The email to be sent to the Store Manager of Store A should be as follows
![image](https://github.com/rodoluca/indicators_automation/assets/115651087/b07c64e3-59e7-446f-a997-572bae03b6ab)


Important Files and Information:

-File "Emails.xlsx" contains the name, store, and email of each manager. Note: I suggest replacing the email column of each manager with your own email for testing purposes.

-File "Vendas.xlsx" contains sales data for all stores. Note: Each manager should only receive the One Page and an attached Excel file with the sales data for their specific store. Information from other stores should not be sent to managers who are not responsible for those stores.

-File "Lojas.csv" contains the name of each store.

-At the end, your routine should also send an email to the management (information is also in the "Emails.xlsx" file) with two rankings attached: a daily ranking and an annual ranking of the stores. Additionally, in the body of the email, highlight the best and worst-performing stores of the day and the best and worst-performing stores of the year. Store rankings are based on revenue.

-Each store's spreadsheets should be saved within the store's folder with the date of the spreadsheet to create a backup history.

One Page Indicators:

-Revenue -> Yearly Target: 1,650,000 / Daily Target: 1,000
-Product Diversity (number of different products sold in that period) -> Yearly Target: 120 / Daily Target: 4
-Average Ticket per Sale -> Yearly Target: 500 / Daily Target: 500
Note: Each indicator should be calculated daily and annually. The daily indicator should be based on the latest available date in the "Vendas.xlsx" spreadsheet.

Note 2: Tip for green and red symbols: Use the characters from this website (https://fsymbols.com/keyboard/windows/alt-codes/list/) and format them using HTML.

