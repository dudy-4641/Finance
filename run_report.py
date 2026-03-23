import streamlit as st
import pandas as pd
import io
import xlsxwriter

st.set_page_config(page_title="Finance Tool", layout="centered")

st.title("📊 מחולל דוחות כספיים")
st.write("העלה את 3 הקבצים כדי לייצר את האקסל המלא עם גיליון הסינון.")

def fix_date_swap(date_val):
    if pd.isna(date_val): return date_val
    d = pd.to_datetime(date_val, errors='coerce')
    if pd.isna(d): return date_val
    if d.month > 2:
        try: return pd.Timestamp(year=d.year, month=d.day, day=d.month)
        except: return d
    return d

uploaded_files = st.file_uploader("בחר קבצים...", accept_multiple_files=True)

if uploaded_files:
    if len(uploaded_files) < 3:
        st.info("ממתין לכל 3 הקבצים...")
    else:
        if st.button("🚀 ייצר דוח אקסל מלא"):
            try:
                all_data = []
                df_mapping = None
                
                for f in uploaded_files:
                    if "budget" in f.name.lower():
                        df_mapping = pd.read_excel(f, skiprows=2)
                        df_mapping.columns = [c.strip() for c in df_mapping.columns]
                        df_mapping['Entity'] = df_mapping['Entity'].str.strip().str.capitalize()
                        df_mapping['Number of account-ERP'] = df_mapping['Number of account-ERP'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                        df_mapping = df_mapping[['Entity', 'Number of account-ERP', 'Budget item']].dropna()

                for f in uploaded_files:
                    if "budget" in f.name.lower(): continue
                    f.seek(0)
                    if f.name.lower().endswith('.csv'):
                        try: content = f.read().decode('utf-8')
                        except: content = f.read().decode('cp1255')
                        df_raw = pd.read_csv(io.StringIO(content))
                    else:
                        df_raw = pd.read_excel(f)

                    is_ltd = "תאריך למאזן" in df_raw.columns or any("תאריך" in str(c) for c in df_raw.columns)
                    
                    if is_ltd:
                        date_col = [c for c in df_raw.columns if "תאריך" in str(c)][0]
                        acc_num = df_raw['חשבון'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                        df_temp = pd.DataFrame({
                            'Entity': 'Ltd',
                            'Date': pd.to_datetime(df_raw[date_col], dayfirst=True, errors='coerce'),
                            'Vendor': df_raw['תאור חשבון נגדי'].fillna('Unknown'),
                            'Account': (acc_num + " - " + df_raw['תאור'].fillna('').astype(str)).str.strip(),
                            'Amount': pd.to_numeric(df_raw['חובה'], errors='coerce').fillna(0) - pd.to_numeric(df_raw['זכות'], errors='coerce').fillna(0),
                            'Memo': df_raw.get('פרטים', '-').fillna('-'),
                            'MapKey': acc_num
                        })
                    else:
                        f.seek(0)
                        df_raw = pd.read_excel(f, skiprows=4) if f.name.endswith(('.xlsx', '.xls')) else pd.read_csv(f, skiprows=4)
                        acc_name = df_raw['Distribution account'].astype(str)
                        acc_num = acc_name.str.extract('(\d+)', expand=False).fillna(acc_name).str.strip()
                        df_temp = pd.DataFrame({
                            'Entity': 'Inc',
                            'Date': df_raw['Transaction date'].apply(fix_date_swap),
                            'Vendor': df_raw['Name'].fillna('Unknown'),
                            'Account': acc_name,
                            'Amount': pd.to_numeric(df_raw['Amount'].astype(str).str.replace(r'[\$,",]', '', regex=True), errors='coerce'),
                            'Memo': df_raw['Memo/Description'].fillna('-'),
                            'MapKey': acc_num
                        })
                    all_data.append(df_temp)

                df_final = pd.concat(all_data, ignore_index=True).dropna(subset=['Date'])
                df_final = pd.merge(df_final, df_mapping, left_on=['Entity', 'MapKey'], right_on=['Entity', 'Number of account-ERP'], how='left')
                df_final['Budget item'] = df_final['Budget item'].fillna('Other / Unmapped')
                
                bs_list = ['Checking', 'Mesh', 'Savings', 'בנק', 'מזומן', 'חו"ז', 'לקוחות', 'ספקים', 'Payable', 'Receivable', 'Credit Card']
                df_final = df_final[~df_final['Account'].str.contains('|'.join(bs_list), na=False, case=False)]

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final[['Entity', 'Date', 'Vendor', 'Account', 'Amount', 'Memo', 'Budget item']].to_excel(writer, sheet_name='Data', index=False)
                    
                    workbook = writer.book
                    ws = workbook.add_worksheet('סינון מאוחד')
                    
                    ents = ["All"] + sorted(df_final['Entity'].unique().tolist())
                    budgs = ["All"] + sorted(df_final['Budget item'].unique().tolist())
                    accs = ["All"] + sorted(df_final['Account'].unique().tolist())
                    months = sorted(df_final['Date'].dt.to_period('M').dt.to_timestamp().unique())

                    ws.write('A1', 'יישות:'); ws.write('C1', 'תקציב:'); ws.write('E1', 'חשבון:'); ws.write('G1', 'מחודש:'); ws.write('I1', 'עד:'); ws.write('K1', 'סה"כ:')
                    
                    ls = workbook.add_worksheet('Lists')
                    for i, v in enumerate(ents): ls.write(i, 0, v)
                    for i, v in enumerate(budgs): ls.write(i, 1, v)
                    for i, v in enumerate(accs): ls.write(i, 2, v)
                    for i, v in enumerate(months): ls.write_datetime(i, 3, v, workbook.add_format({'num_format': 'mm/yyyy'}))

                    ws.data_validation('B1', {'validate': 'list', 'source': f'=Lists!$A$1:$A${len(ents)}'})
                    ws.data_validation('D1', {'validate': 'list', 'source': f'=Lists!$B$1:$B${len(budgs)}'})
                    ws.data_validation('F1', {'validate': 'list', 'source': f'=Lists!$C$1:$C${len(accs)}'})
                    ws.data_validation('H1', {'validate': 'list', 'source': f'=Lists!$D$1:$D${len(months)}'})
                    ws.data_validation('J1', {'validate': 'list', 'source': f'=Lists!$D$1:$D${len(months)}'})

                    ws.write('B1', 'All'); ws.write('D1', 'All'); ws.write('F1', 'All')
                    if months:
                        ws.write_datetime('H1', months[0])
                        ws.write_datetime('J1', months[-1])

                    h_fmt = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
                    headers = ['Entity', 'Date', 'Vendor', 'Account', 'Amount', 'Memo', 'Budget Item']
                    for i, h in enumerate(headers):
                        ws.write(3, i, h, h_fmt)
                        ws.set_column(i, i, 18)

                    last_r = len(df_final) + 1
                    cond = f'(IF($B$1="All", 1, Data!$A$2:$A${last_r}=$B$1)) * (IF($D$1="All", 1, Data!$G$2:$G${last_r}=$D$1)) * (IF($F$1="All", 1, Data!$D$2:$D${last_r}=$F$1)) * (Data!$B$2:$B${last_r}>=$H$1) * (Data!$B$2:$B${last_r}<=EOMONTH($J$1,0))'
                    
                    ws.write_dynamic_array_formula('A5:A5', f'=IFERROR(FILTER(Data!A2:G{last_r}, {cond}), "אין נתונים")')
                    ws.write_formula('L1', '=SUM(E5:E20000)', workbook.add_format({'num_format': '#,##0.00', 'bold': True, 'bg_color': '#FFEB9C'}))

                st.success("✅ הדוח המלא מוכן!")
                st.download_button(label="📥 הורד אקסל עם סינון", data=output.getvalue(), file_name="Finance_Report_Full.xlsx")
            except Exception as e:
                st.error(f"שגיאה: {e}")
