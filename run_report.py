import streamlit as st
import pandas as pd
import io
import xlsxwriter

st.set_page_config(page_title="Finance Tool", layout="wide")

st.title("📊 מחולל דוחות כספיים")
st.write("העלה את 3 הקבצים כדי לייצר את האקסל המלא.")

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
        st.info("ממתין ל-3 הקבצים (LTD, INC, Budget)...")
    else:
        if st.button("🚀 לייצר דוח אקסל"):
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
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final[['Entity', 'Date', 'Vendor', 'Account', 'Amount', 'Memo', 'Budget item']].to_excel(writer, sheet_name='Data', index=False)
                    workbook = writer.book
                    ws = workbook.add_worksheet('Summary')
                    ws.write('A1', 'הדוח מוכן בגיליון Data')
                    ws.write_formula('B2', '=SUM(Data!E:E)')

                st.success("✅ הצלחנו!")
                st.download_button(label="📥 הורד אקסל", data=output.getvalue(), file_name="Report.xlsx")
            except Exception as e:
                st.error(f"שגיאה: {e}")
