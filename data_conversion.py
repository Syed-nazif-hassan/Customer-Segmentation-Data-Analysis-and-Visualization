import pandas

# Read the text file
text_file = 'Ecommerce Purchases.txt'
df = pandas.read_csv(text_file, delimiter=',')

# Save as Excel file
excel_file = 'ecommerce_purchases.xlsx'
df.to_excel(excel_file, index=False)