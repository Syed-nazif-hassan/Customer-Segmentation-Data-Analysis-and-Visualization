import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import os


# Load the dataset
df = pd.read_excel('ecommerce_purchases.xlsx')

# Segmentation by Spending
threshold = df['Purchase Price'].mean()

df['Spender Type'] = df['Purchase Price'].apply(
    lambda x: 'High Spender' if x >= threshold else 'Low Spender')

plt.figure(figsize=(8, 6))
ax = sns.countplot(data=df, x='Spender Type',
                   hue='Spender Type', palette='Set2', legend=False)
plt.title('Number of High Spenders vs Low Spenders')
plt.xlabel('Spender Type')
plt.ylabel('Count')

for p in ax.patches:
    count = int(p.get_height())
    ax.annotate(f'{count}', (p.get_x() + p.get_width() / 2., p.get_height()),
                ha='center', va='baseline', fontsize=12, color='black', xytext=(0, 5),
                textcoords='offset points')

plt.ylim(0, ax.get_ylim()[1] * 1.1)
plt.grid(axis='y', color='black')
plt.gca().set_facecolor('lightgrey')
sns.despine()

plt.show()

# Segmentation by Demographic
# Segmentation by Job
job_spender_type_counts = df.groupby(
    ['Job', 'Spender Type']).size().unstack(fill_value=0)

folder_path = os.path.join(os.getcwd(), 'custom_excel_file')
file_path = os.path.join(folder_path, 'custom.xlsx')

os.makedirs(folder_path, exist_ok=True)
job_spender_type_counts.to_excel(
    file_path, sheet_name='Job_Spender_Type_Counts', index=True)

# Segmentation by Company
company_spender_type_counts = df.groupby(
    ['Company', 'Spender Type']).size().unstack(fill_value=0)

with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    company_spender_type_counts.to_excel(
        writer, sheet_name='Company_Spender_Type_Counts', index=True)

# Segmentation by Language
Languages = df['Language'].value_counts()

language_dict = {
    'de': 'German',
    'ru': 'Russian',
    'el': 'Greek',
    'pt': 'Portuguese',
    'en': 'English',
    'fr': 'French',
    'es': 'Spanish',
    'it': 'Italian',
    'zh': 'Chinese'
}
language_df = pd.DataFrame({
    'Language Short': Languages.index,
    'Count': Languages.values
})

language_df['Language Long'] = language_df['Language Short'].map(language_dict)

plt.figure(figsize=(8, 6))
plt.title('Languages of Customers')
plt.pie(language_df['Count'], labels=language_df['Language Short'] + ' - ' + language_df['Language Long'],
        autopct='%1.1f%%', startangle=140, colors=sns.color_palette('pastel', len(language_df)))

plt.show()

# Behavioral Segmentation
am_pm_purchase_stats = df.groupby(
    'AM or PM')['Purchase Price'].agg(['size', 'mean', 'sum'])

# Mean Purchase for AM and PM
purchase_price_mean = am_pm_purchase_stats['mean'].values

print(f'Mean purchase price for AM: {round(
    purchase_price_mean[0], 4)}\nMean purchase price for PM: {round(purchase_price_mean[1], 4)}')

# Total Purchase Amount by AM and PM
plt.figure(figsize=(8, 6))
sns.lineplot(data=am_pm_purchase_stats, x=am_pm_purchase_stats.index,
             y='sum', marker='o', color='brown', label='Total Purchase Amount')

for index, row in am_pm_purchase_stats.iterrows():
    plt.text(index, row['sum'], f'{
             row["sum"]:.2f}', ha='center', va='bottom', fontsize=10, color='black')

plt.title('Total Purchase Amount by AM and PM')
plt.xlabel('AM and PM')
plt.ylabel('Total Sum of Purchase Price')
plt.ylim(0, am_pm_purchase_stats['sum'].max() * 1.1)
plt.grid(True, color='black')
plt.legend()
sns.despine()
plt.gca().set_facecolor('lightgrey')

plt.show()

# Total Number of Purchases by AM and PM
plt.figure(figsize=(8, 6))
ax = sns.barplot(data=am_pm_purchase_stats, x=am_pm_purchase_stats.index,
                 hue=am_pm_purchase_stats.index, y='size', palette='Accent')
plt.title('Total Number of Purchases by AM and PM')
plt.xlabel('AM and PM')
plt.ylabel('Total Number of Purchases')

for p in ax.patches:
    count = int(p.get_height())
    ax.annotate(f'{count}', (p.get_x() + p.get_width() / 2., p.get_height()),
                ha='center', va='baseline', fontsize=12, color='black', xytext=(0, 5),
                textcoords='offset points')

plt.ylim(0, ax.get_ylim()[1] * 1.1)
plt.grid(axis='y', color='black')
plt.gca().set_facecolor('lightgrey')
sns.despine()

plt.show()

# Technographic Segmentation
mozilla_users = df[df['Browser Info'].str.contains('Mozilla', case=False)]
mozilla_user_info = mozilla_users[['IP Address', 'Email']]

with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    mozilla_user_info.to_excel(
        writer, sheet_name='Mozilla_User_Info', index=False)

# New Custom Excel File
custom_ecommerce_purchases_file_path = os.path.join(
    folder_path, 'custom_ecommerce_purchases.xlsx')

if not os.path.exists(custom_ecommerce_purchases_file_path):
    df.to_excel(custom_ecommerce_purchases_file_path,
                sheet_name='Custom_Ecommerce_Purchases', index=False)
else:
    with pd.ExcelWriter(custom_ecommerce_purchases_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(
            writer, sheet_name='Custom_Ecommerce_Purchases', index=False)
