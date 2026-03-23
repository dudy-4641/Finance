import pandas as pd
import os
import glob

# 1. איתור הקובץ - הסקריפט יחפש כל קובץ CSV בתיקייה
csv_files = glob.glob("*.csv")
if not csv_files:
    print("לא נמצא קובץ CSV בתיקייה. וודא שקובץ ה-QuickBooks נמצא באותה תיקייה של הסקריפט.")
    input("לחץ על Enter ליציאה...")
    exit()

file_name = csv_files[0] # לוקח את הקובץ הראשון שנמצא
print(f"מעבד את הקובץ: {file_name}")

# 2. קריאה ועיבוד
df = pd.read_csv(file_name, skiprows=4)
df = df.dropna(subset=['Transaction date'])
df['Amount'] = df['Amount'].astype(str).str.replace(',', '').str.replace('"', '').astype(float)
df['Name'] = df['Name'].fillna('Unknown')
df['Memo/Description'] = df['Memo/Description'].fillna('-')
df['Distribution account'] = df['Distribution account'].fillna('Uncategorized')

# 3. סינון כרטיסים תוצאתיים בלבד
balance_sheet_terms = ['Checking', 'Mesh', 'Morgan Stanley', 'Savings', 'Money Market', 'Investments', 'Intercompany', 'Payable', 'Receivable', 'Accrued', 'Prepaid', 'Stripe', 'Balance', 'Cash', 'Credit Card']
mask = df['Distribution account'].str.contains('|'.join(balance_sheet_terms), na=False, case=False)
df_result = df[~mask].copy()
df_result = df_result[['Transaction date', 'Name', 'Distribution account', 'Amount', 'Transaction type', 'Memo/Description']]
df_result = df_result.sort_values(by=['Distribution account', 'Transaction date'])

# 4. יצירת האקסל החכם
output_file = 'QB_Smart_Report_Automated.xlsx'
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    df_result.to_excel(writer, sheet_name='Data', index=False)
    workbook = writer.book
    filter_sheet = workbook.add_worksheet('כלי סינון חכם')
    
    # עיצובים
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1, 'align': 'center'})
    
    # ממשק סינון
    filter_sheet.write('A1', 'בחר חשבון:', header_fmt)
    unique_accs = sorted(df_result['Distribution account'].unique())
    list_sheet = workbook.add_worksheet('AccList')
    for i, acc in enumerate(unique_accs):
        list_sheet.write(i, 0, acc)
    
    filter_sheet.data_validation('B1', {'validate': 'list', 'source': f'=AccList!$A$1:$A${len(unique_accs)}'})
    if unique_accs: filter_sheet.write('B1', unique_accs[0])
    
    headers = ['תאריך', 'ספק', 'חשבון', 'סכום', 'סוג', 'תיאור (Memo)']
    for col, h in enumerate(headers):
        filter_sheet.write(3, col, h, header_fmt)
        filter_sheet.set_column(col, col, 18)

    last_row = len(df_result) + 1
    formula = f'=FILTER(Data!A2:F{last_row}, Data!C2:C{last_row} = B1, "אין נתונים")'
    filter_sheet.write_dynamic_array_formula('A5:F5', formula)

print(f"הצלחתי! הקובץ {output_file} נוצר בתיקייה.")
