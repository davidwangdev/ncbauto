import pandas as pd
from django.shortcuts import render
from django.http import HttpResponse
from .forms import UploadFileForm
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

def home(request):
    return render(request, 'fileprocessor/home.html')

def charges(request):
    if request.method == 'POST':
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            uploaded_file = request.FILES['file']
            output = handle_charges(uploaded_file)
            response = HttpResponse(
                output.getvalue(),
                content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            response['Content-Disposition'] = 'attachment; filename=Charges_Summary.xlsx'
            return response
    else:
        form = UploadFileForm()
    return render(request, 'fileprocessor/charges.html', {'form': form})

def surgeries(request):
    if request.method == 'POST':
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            uploaded_file = request.FILES['file']
            output = handle_surgeries(uploaded_file)
            response = HttpResponse(
                output.getvalue(),
                content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            response['Content-Disposition'] = 'attachment; filename=Surgeries_Summary.xlsx'
            return response
    else:
        form = UploadFileForm()
    return render(request, 'fileprocessor/surgeries.html', {'form': form})

def handle_charges(file):
    # Read main sheet, calculate aggregate sum
    if(file.name.endswith('.xls') | file.name.endswith('.xlsx')):
        df = pd.read_excel(file)

    df['SUB GL'] = 'SUB GL ' + df['SUB GL'].astype(str)
    aggregate_sum = df.groupby('SUB GL')['ExtCost'].sum().reset_index()
    total_sum = aggregate_sum['ExtCost'].sum().round(2)
    aggregate_sum = pd.concat([pd.DataFrame({'SUB GL': ['Total Ext Cost'], 'ExtCost': [total_sum]}), aggregate_sum])
    aggregate_sum['Source'] = 'Aggregate'

    # Sheets to be processed (assumes sheet names will remain same over time)
    sheets = ['CIV', 'RESOLUTE', 'SLB', 'MTB', 'NEB', 'NCB', 'BMC']
    all_data = [aggregate_sum]

    # Process each sheet and calculate sum per 'SUB GL'
    for sheet in sheets:
        df_sheet = pd.read_excel(file, sheet_name=sheet)
        df_sheet['SUB GL'] = 'SUB GL ' + df_sheet['SUB GL'].astype(str)
        sum_sheet = df_sheet.groupby('SUB GL')['ExtCost'].sum().reset_index()
        total_sum = sum_sheet['ExtCost'].sum().round(2)
        sum_sheet = pd.concat([pd.DataFrame({'SUB GL': ['Total Ext Cost'], 'ExtCost': [total_sum]}), sum_sheet])
        sum_sheet['Facility'] = sheet
        all_data.append(sum_sheet)

    # Convert 'ExtCost' columns to accounting format
    for df in all_data:
        columns_to_format = [col for col in df.columns if 'ExtCost' in col]
        for col in columns_to_format:
            df[col] = df[col].apply(lambda x: f'${x:,.2f}' if pd.notnull(x) else x)

    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"

    # Styling definitions
    header_fill = PatternFill(start_color="9BC2E6", end_color="9BC2E6", fill_type="solid")
    cell_fill = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
    font = Font(bold=True, color="000000")
    alignment_header = Alignment(horizontal="center")
    alignment_accounting = Alignment(horizontal="right")

    # Write each dataframe to Excel in different columns
    start_col = 1
    for df in all_data:
        # Write each header and style header
        for col_num, column_title in enumerate(df.columns, start=start_col):
            cell = ws.cell(row=1, column=col_num, value=column_title)
            cell.fill = header_fill
            cell.font = font
            cell.alignment = alignment_header

        for row_num, row in enumerate(dataframe_to_rows(df, index=False, header=False), start=2):
            for col_num, cell_value in enumerate(row, start=start_col):
                cell = ws.cell(row=row_num, column=col_num, value=cell_value)
                if df.columns[col_num - start_col] in ['SUB GL', 'ExtCost', 'Facility']:
                    cell.fill = cell_fill
                    if df.columns[col_num - start_col] in columns_to_format:
                        cell.alignment = alignment_accounting

        # Auto adjust column width to ensure readability
        for col_num, column_title in enumerate(df.columns, start=start_col):
            max_length = max(
                len(str(column_title)),
                *(len(str(cell_value)) for cell_value in df[column_title])
            )
            ws.column_dimensions[ws.cell(row=1, column=col_num).column_letter].width = max_length + 2

        # Add break in columns
        start_col += len(df.columns) + 1

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def handle_surgeries(file):
    if file.name.endswith('.csv'):
        df = pd.read_csv(file)
    elif file.name.endswith(('.xls', '.xlsx')):
        df = pd.read_excel(file)
    else:
        raise ValueError("Unsupported file format. Only CSV and XLS/XLSX are supported.")

    surgeon_names = [
        "BURGESS MD, MEGAN", "CYR MD, STEVEN", "DEBERARDINO MD, THOMAS", 
        "FERGUSON MD, EARL", "GARCIA MD, FRANCISCO", "KAISER MD, BRYAN", 
        "KREINES DO, ALEXANDER", "LYNCH MD, JAMIE", "NILSSON MD, JOEL", 
        "SWANN MD, MATTHEW", "VIROSLAV MD, SERGIO"
    ]

    multiply_by_4 = ["DEBERARDINO", "KREINES", "NILSSON", "LYNCH"]

    # Count occurrences of each surgeon in dataframe
    surgeon_count = {}
    for surgeon_name in surgeon_names:
        last_name = surgeon_name.split(' ')[0]  # Get last name
        count = df[df['SURGEON'].str.contains(surgeon_name, na=False)].shape[0]
        total_bags = count * 4 if last_name in multiply_by_4 else count
        surgeon_count[last_name] = (count, total_bags)

    print(surgeon_count)

    # Write results to Excel file
    wb = Workbook()
    ws = wb.active
    ws.title = "Surgeon Counts"

    # Styling definitions
    header_fill = PatternFill(start_color="9BC2E6", end_color="9BC2E6", fill_type="solid")
    font = Font(bold=True, color="000000")
    alignment_header = Alignment(horizontal="center")

    # Write headers
    ws.append(['Surgeon Last Name', 'Count', 'Total Bags', 'Bag Name'])
    header_row = ws[1]
    for cell in header_row:
        cell.fill = header_fill
        cell.font = font
        cell.alignment = alignment_header

    # Write data
    for last_name, count in surgeon_count.items():
        if(last_name == "BURGESS"):
            original_count, total_bags = count
            ws.append([last_name, original_count, total_bags, 'Tumescent Solution – NS 1000ml + TXA 1gm + Lidocaine 1% 50ml + Epi 1ml. Profile but wait for request to mix due to low stability.'])
            continue
        if(last_name == "CYR"):
            original_count, total_bags = count
            ws.append([last_name, original_count, total_bags, 'Heparin 30,000 unit in 1000ml NS for Cell Saver.'])
            continue
        if(last_name == "DEBERARDINO"):
            original_count, total_bags = count
            ws.append([last_name, original_count, total_bags, '1mg Epinephrine in NS 3000ml irrigation'])
            continue
        original_count, total_bags = count
        ws.append([last_name, original_count, total_bags])

    # Auto adjust column width
    for col in ws.columns:
        max_length = 0
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2) * 1.2
        ws.column_dimensions[col[0].column_letter].width = adjusted_width

    # Save the workbook to a BytesIO object
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output
