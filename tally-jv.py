from flask import Flask, render_template, request, send_file
import pandas as pd
import re
from datetime import datetime
import os

app = Flask(__name__)

# Function to process Excel file
def process_excel(file):
    df = pd.read_excel(file)

    df1 = df[['UID', 'BudgetItem', 'Department', 'Amount', 'Vendor/Transfer to department', 'AmountRemarks']]
    df2 = df1.rename(columns={
        'UID': 'Entry No',
        'BudgetItem': 'Dr Ledger Name',
        'Department': 'Dr Cost Center',
        'Amount': 'Dr Amt',
        'Vendor/Transfer to department': 'Cr Ledger Name'
    })

    def extract_reference(text):
        bill_match = re.search(r'Bill No\.\s*(\S+)', str(text))
        invoice_match = re.search(r'Invoice No\.\s*(\S+)', str(text))
        quotation_match = re.search(r'Quotation Ref No\.\s*(\S+)', str(text))
        if bill_match:
            return bill_match.group(1)
        elif invoice_match:
            return invoice_match.group(1)
        elif quotation_match:
            return quotation_match.group(1)
        return None

    df2['Bill Ref No.'] = df2['AmountRemarks'].apply(extract_reference)
    df2['Vch Narration'] = 'UID No. ' + df2['Entry No'].astype(str) + ' ' + df2['AmountRemarks']

    df3 = df2[['Entry No', 'Dr Ledger Name', 'Dr Cost Center', 'Dr Amt', 'Cr Ledger Name', 'Bill Ref No.', 'Vch Narration']]
    df3.insert(0, 'Date', '')
    df3.insert(2, 'Vch Name', '')
    df3.insert(7, 'Cr Cost Center', '')
    df3.insert(8, 'Cr Amt', '')

    df3['Dr Amt'] = df3['Dr Amt'].astype(str).str.replace(',', '', regex=True)
    df3['Dr Amt'] = pd.to_numeric(df3['Dr Amt'])
    df4 = df3.copy()
    df4['Cr Amt'] = df3.groupby('Entry No')['Dr Amt'].transform('sum')

    df4['Dr Amt'] = df4['Dr Amt'].apply(lambda x: f"{x:,.2f}")
    df4['Cr Amt'] = df4['Cr Amt'].apply(lambda x: f"{x:,.2f}")

    # ----- Your newly requested logic starts here -----
    df5 = df4.drop_duplicates()
    df5 = df5[~df5.apply(lambda row: row.astype(str).str.contains('Shri Chaitanya Health and Care Trust \(Branch Transfer\)').any(), axis=1)]

    # Replace blank or NaN values in 'Bill Ref No.' with values from 'Entry No'
    df5['Bill Ref No.'] = df5['Bill Ref No.'].replace('', pd.NA)  # convert empty strings to NA
    df5['Bill Ref No.'] = df5['Bill Ref No.'].fillna(df5['Entry No'])

    # Remove "Internal Transfer" narration if still needed
    df5 = df5[~df5['Vch Narration'].str.contains("Internal Transfer", na=False)]
    # ----- Your newly requested logic ends here -----

    # Fill 'Date' and 'Vch Name'
    today = datetime.today().strftime('%d/%m/%Y')
    df5['Date'] = today
    df5['Vch Name'] = 'Journal'

    # Final column order
    final_df = df5[['Date', 'Entry No', 'Vch Name', 'Dr Ledger Name', 'Dr Amt', 'Dr Cost Center',
                    'Cr Ledger Name', 'Cr Amt', 'Cr Cost Center', 'Vch Narration', 'Bill Ref No.']]

    output_filename = f'ICT_to_TDL_Format_{datetime.now().strftime("%Y-%m-%d_%H-%M-%S")}.xlsx'
    final_df.to_excel(output_filename, index=False)

    return output_filename

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return "No file uploaded", 400

    file = request.files['file']
    if file.filename == '':
        return "No selected file", 400

    output_file = process_excel(file)
    return send_file(output_file, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
