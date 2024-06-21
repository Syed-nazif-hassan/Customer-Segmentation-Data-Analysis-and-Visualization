import pandas as pd

# Load the dataset
df = pd.read_excel('ecommerce_purchases.xlsx')

print(df)

print(df.info())

# Pandas describe function in chunks
num_cols = len(df.columns)
chunk_size = 4

for i in range(0, num_cols, chunk_size):
  group_cols = df.columns[i:i+chunk_size]
  print(df[group_cols].describe(include='all'))
  print("\n")
  
print(df.isnull().sum())
print(df.isna().any().any())
