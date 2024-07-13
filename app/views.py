import pandas as pd
from django.shortcuts import render
from django.http import HttpResponse
from django.contrib import messages
from io import BytesIO

def convert(request):
    if request.method == 'POST':
        messages.info(request, "Conversion process has started.")
        # Read the Excel file into DataFrames
        excel_file_1 = request.FILES['file1']
        excel_file_2 = request.FILES['file2']
        excel_file_3 = request.FILES['file3']
        excel_file_4 = request.FILES['file4']

        
        df1 = pd.read_excel(excel_file_1) # Billing
        df2 = pd.read_excel(excel_file_2) # Auto
        df3 = pd.read_excel(excel_file_3) # SPU
        df4 = pd.read_excel(excel_file_4) # SCS

        # Define extraction functions
        def df1_sheet_extract(row):
            # Check for NaN in 'Sales Document' and 'Material'
            if pd.isna(row['Sales Document']) or pd.isna(row['Material']):
                return 'None'

            # Convert 'Sales Document' to int and concatenate with 'Material'
            link = int(row['Sales Document'])
            link = str(link) + str(row['Material'])

            # Remove any periods from the string
            return link.replace('.', '')
        def df2_sheet_extract(row):
            indent = str(row['Indent'])  # Remove leading zeros
            link = indent + str(row['Spare Sap Code'])
            return link

        def df3_sheet_extract(row):
            ticket = str(row['Ticket'])
            item_code = str(row['Item Code'])
            link = ticket + item_code
            return link

        def df4_sheet_extract(row):
            ticket = str(row['Ticket No']).replace('.0', '')
            item_code = str(row['Spare'])
            link = ticket + item_code
            return link

        # Apply the functions to create 'Link' columns
        df1['Link'] = df1.apply(df1_sheet_extract, axis=1)
        df2['Link'] = df2.apply(df2_sheet_extract, axis=1)
        df3['Link2'] = df3.apply(df3_sheet_extract, axis=1)
        df4['Link2'] = df4.apply(df4_sheet_extract, axis=1)


        # Merging DataFrames
        f3 = df1[[
            'Purchase order number', 'Billing Date', 'Billing Document', 'Link',
            'Party Name',  # Rename to Party Code
            'Material',
            'Material Description',
            'Billed Quantity',
            'Gross Value before TP/Wrty Support',
            'Sales Document',
        ]].merge(df2[[
            'Link', 'Ticket ID', 
            'Model',  # Rename to Machine
            'Machine Status',
            'Product',
            'Frcode',
            'Franchise Name',
            'Indent',
            'Spare Sap Code',
            'Spare Part Description',
            'SO Quantity', 'CreatedDate'
        ]], on="Link", how="left")

        def f3_sheet_extract(row):
            ticket = str(row['Ticket ID']).replace('.0', '')
            item_code = str(row['Material'])
            link = ticket + item_code
            return link 

        f3['Link2'] = f3.apply(f3_sheet_extract, axis=1)
        f3.rename(columns={'Model': 'Machine', }, inplace=True)

        f4 = f3[[
            'Purchase order number', 'Billing Date', 'Billing Document', 'Link2',
            'Party Name',
            'Product',
            'Frcode',
            'Franchise Name',
            'Material',
            'Material Description',
            'Billed Quantity',
            'Gross Value before TP/Wrty Support',
            'Sales Document',
            'Ticket ID', 
            'Machine',  # Rename to Machine
            'Machine Status',
            'Indent',
            'Spare Sap Code',
            'Spare Part Description',
            'SO Quantity', 'CreatedDate'
        ]].merge(df3[[
            'Link2',
            'SPU NO',
            'po ref no',  # 'po ref no': 'Credit Number under SPU'
        ]], on="Link2", how="left")

        f4.rename(columns={'po ref no': 'Credit Number under SPU'}, inplace=True)

        f5 = f4[[
            'Purchase order number', 'Billing Date', 'Billing Document',
            'Party Name',
            'Product',
            'Frcode',
            'Franchise Name',
            'Material',
            'Material Description',
            'Billed Quantity',
            'Gross Value before TP/Wrty Support',
            'Sales Document',
            'Ticket ID', 
            'Machine',  # Rename to Machine
            'Machine Status',
            'Link2',
            'SPU NO',
            'Credit Number under SPU',  # 'po ref no': 'Credit Number under SPU'.
            'Indent',
            'Spare Sap Code',
            'Spare Part Description',
            'SO Quantity', 'CreatedDate'
        ]].merge(df4[[
            'Link2','CustomerName','Technician'
        ]], on="Link2", how="left")

        # Define a function to handle empty 'Billing Document'
        def handle_billing_doc(row):
            if pd.isna(row['Billing Document']) or row['Billing Document'] == '':
                return f"{row['Material']} - {row['Material Description']}"
            return row['Billing Document']

        # Apply the function to the 'Billing Document' column
        f5['New Billing Doc'] = f5.apply(handle_billing_doc, axis=1)

        f5_selected = pd.DataFrame({
            'Indent': f5['Indent'],
            'Spare Sap Code': f5['Spare Sap Code'],
            'Spare Part Description': f5['Spare Part Description'],
            'SO Quantity': f5['SO Quantity'],
            'Ticket ID': f5['Ticket ID'],
            'Machine Status': f5['Machine Status'],
            'Product': f5['Product'],
            'Model': f5['Machine'],
            'Frcode': f5['Frcode'],
            'Franchise Name': f5['Franchise Name'],
            'Billing Document': f5['New Billing Doc'],
            'Date': f5['CreatedDate'],


        })
        f5_selected.rename(columns={
            'Machine': 'Model',
            'New Billing Doc': 'Billing Document'
        }, inplace=True)

        f5 = f5.drop(columns=['Product', 'Frcode', 'Franchise Name', 'New Billing Doc', 'Indent', 'Spare Sap Code', 'Spare Part Description','SO Quantity', 'CreatedDate'])

        # Create a BytesIO object to save the Excel file to memory
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            f5.to_excel(writer, index=False, sheet_name='Output 1')
            f5_selected.to_excel(writer, sheet_name='Pending Call', index=False)

            workbook = writer.book
            worksheet = writer.sheets['Output 1']
            worksheet2 = writer.sheets['Pending Call']

            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': False,
                'valign': 'top',
                'fg_color': '#6699cb',  # Light blue color
                'border': 1,  # Bold border
                'align': 'left',  # Center align text
                'font_color': '#FFFFFF'  # White font color,
            })
            cell_format = workbook.add_format({
                'align': 'left',  # Left align text
            })
            # Format the first sheet (Output 1)
            for i, col in enumerate(f5.columns):
                column_len = max(f5[col].astype(str).map(len).max(), len(col)) + 1  # Adjust for border width
                worksheet.set_column(i, i, column_len, cell_format)
                worksheet.write(0, i, col, header_format)  # Apply the header format

            # Format the second sheet (Pending Call)
            for i, col in enumerate(f5_selected.columns):
                column_len = max(f5_selected[col].astype(str).map(len).max(), len(col)) + 1  # Adjust for border width
                worksheet2.set_column(i, i, column_len, cell_format)
                worksheet2.write(0, i, col, header_format)  # Apply the header format

        # Seek to the beginning of the stream
        output.seek(0)

        # Create the HTTP response
        response = HttpResponse(output, content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        response['Content-Disposition'] = 'attachment; filename=Results.xlsx'
        
        messages.success(request, "Conversion process completed successfully.")
        return response

        return render(request, 'index.html')

    return render(request, 'index.html')
