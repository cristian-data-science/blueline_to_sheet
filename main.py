import pandas as pd
import pymssql
import var


conn = pymssql.connect(var.server, var.username, var.password, var.database)
cursor = conn.cursor(as_dict=True)

cursor.execute("SELECT DISTINCT(SALESORDERNUMBER) AS SALESORDERNUMBER, DELIVERYADDRESSNAME, CUSTOMERREQUISITIONNUMBER as OC, CUSTOMSDOCUMENTDATE FROM SalesOrderLineV2Staging WHERE CUSTOMERREQUISITIONNUMBER <> '' AND CUSTOMSDOCUMENTDATE BETWEEN GETDATE() -30  and GETDATE() ")
df_sales_order = pd.DataFrame(cursor.fetchall())
#df_sales_order

cursor.execute("SELECT SALESORDERNUMBER, INVOICENUMBER, INVOICEDATE FROM SalesInvoiceHeaderV2Staging WHERE INVOICEDATE BETWEEN GETDATE() -30  and GETDATE() ")
df_sales_invoice = pd.DataFrame(cursor.fetchall())

df_sql = df_sales_invoice.merge(df_sales_order, how='left', on='SALESORDERNUMBER')
df_sql = df_sql.dropna()

df_excel = pd.read_excel("nuevo_Rep.xlsx")
df_final = df_sql.merge(df_excel, how='right', on='INVOICENUMBER')

df_final = df_final.dropna()
df_final.to_excel("finalfinal.xlsx")
cursor.close()
conn.close()