import pandas as pd
import xlsxwriter

def col_n(n):
    """Convert a number to its corresponding Excel column letter."""
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result

# Historical Data
years_historical = [2021, 2022, 2023]
revenue_historical = [85000, 90000, 95000]
costs_historical = [40000, 42000, 44000]

# User Inputs for Projections
num_projection_years = 10
growth_rate = 0.05  # This will be set dynamically in the Excel Control Panel

# Calculating Projection Years
years_projection = [years_historical[-1] + i + 1 for i in range(num_projection_years)]

# Income Statement Data
income_statement = {
    'Metric': ['Revenue', 'Costs', 'Profit'],
    **{year: [revenue_historical[i] if i < len(years_historical) else None,
              costs_historical[i] if i < len(years_historical) else None,
              revenue_historical[i] - costs_historical[i] if i < len(years_historical) else None
             ] for i, year in enumerate(years_historical + years_projection)}
}

# Balance Sheet Data
assets_historical = [150000, 160000, 170000]
liabilities_historical = [70000, 75000, 80000]
equity_historical = [80000, 85000, 90000]

balance_sheet = {
    'Metric': ['Assets', 'Liabilities', 'Equity'],
    **{year: [assets_historical[i] if i < len(years_historical) else None,
              liabilities_historical[i] if i < len(years_historical) else None,
              equity_historical[i] if i < len(years_historical) else None
             ] for i, year in enumerate(years_historical + years_projection)}
}

# Cash Flow Statement Data
operating_cf_historical = [20000, 21000, 22000]
investing_cf_historical = [-10000, -10500, -11000]
financing_cf_historical = [15000, 16000, 17000]

cash_flow_statement = {
    'Metric': ['Operating Cash Flow', 'Investing Cash Flow', 'Financing Cash Flow'],
    **{year: [operating_cf_historical[i] if i < len(years_historical) else None,
              investing_cf_historical[i] if i < len(years_historical) else None,
              financing_cf_historical[i] if i < len(years_historical) else None
             ] for i, year in enumerate(years_historical + years_projection)}
}

# Write DataFrames to Excel with formatting
with pd.ExcelWriter('financial_model.xlsx', engine='xlsxwriter') as writer:
    # Create Control Panel Sheet
    workbook = writer.book
    worksheet_control = workbook.add_worksheet('Control Panel')
    worksheet_control.write('A1', 'Growth Rate')
    worksheet_control.write('B1', growth_rate)
    worksheet_control.data_validation('B1', {'validate': 'decimal', 'criteria': 'between', 'minimum': 0, 'maximum': 1})

    # Convert dictionaries to DataFrames
    df_income_statement = pd.DataFrame(income_statement)
    df_balance_sheet = pd.DataFrame(balance_sheet)
    df_cash_flow_statement = pd.DataFrame(cash_flow_statement)

    # Write all DataFrames to the same sheet
    worksheet = workbook.add_worksheet('Financial Statements')

    # Remove gridlines from the sheet
    worksheet.hide_gridlines(2)  # 0 = show gridlines, 1 = hide printed gridlines, 2 = hide all gridlines

    df_income_statement.to_excel(writer, sheet_name='Financial Statements', startrow=0, startcol=1, index=False)
    df_balance_sheet.to_excel(writer, sheet_name='Financial Statements', startrow=len(df_income_statement) + 2, startcol=1, index=False)
    df_cash_flow_statement.to_excel(writer, sheet_name='Financial Statements', startrow=len(df_income_statement) + len(df_balance_sheet) + 4, startcol=1, index=False)

    # Formatting
    currency_format = workbook.add_format({'num_format': '#,##0'})
    header_format = workbook.add_format({'bold': True, 'font_color': 'blue'})
    historical_format = workbook.add_format({'bg_color': '#D3D3D3', 'num_format': '#,##0'})  # Light Gray Background
    projection_format = workbook.add_format({'num_format': '#,##0'})

    worksheet.set_column('B:Z', 18, currency_format)
    worksheet.write('B1', 'Income Statement', header_format)
    worksheet.write('B{}'.format(len(df_income_statement) + 3), 'Balance Sheet', header_format)
    worksheet.write('B{}'.format(len(df_income_statement) + len(df_balance_sheet) + 5), 'Cash Flow Statement', header_format)

    # Apply Formatting for Historical Data
    for col, year in enumerate(years_historical, start=2):
        worksheet.write_column(1, col, df_income_statement[year], historical_format)
        worksheet.write_column(len(df_income_statement) + 3, col, df_balance_sheet[year], historical_format)
        worksheet.write_column(len(df_income_statement) + len(df_balance_sheet) + 5, col, df_cash_flow_statement[year], historical_format)

    # Apply Formulas for Projections
    for col, year in enumerate(years_projection, start=len(years_historical) + 2):
        worksheet.write_formula(1, col, f'='+col_n(col)+f'{2}*(1+\'Control Panel\'!$B$1)', projection_format)
        worksheet.write_formula(2, col, f'='+col_n(col)+f'{3}*(1+\'Control Panel\'!$B$1)', projection_format)
        worksheet.write_formula(3, col, f'='+col_n(col+1)+f'{2}-'+col_n(col+1)+f'{3}', projection_format)

        worksheet.write_formula(len(df_income_statement) , col+ 3, f'=C{col-1}*(1+\'Control Panel\'!$B$1)', projection_format)
        worksheet.write_formula(len(df_income_statement) , col+ 4, f'=C{col-1}*(1+\'Control Panel\'!$B$1)', projection_format)
        worksheet.write_formula(len(df_income_statement) , col+ 5, f'=B{col+1}-C{col+1}', projection_format)

        worksheet.write_formula(len(df_income_statement) + len(df_balance_sheet) + 5, col, f'=C{col-1}*(1+\'Control Panel\'!$B$1)', projection_format)
        worksheet.write_formula(len(df_income_statement) + len(df_balance_sheet) + 6, col, f'=C{col-1}*(1+\'Control Panel\'!$B$1)', projection_format)
        worksheet.write_formula(len(df_income_statement) + len(df_balance_sheet) + 7, col, f'=B{col+1}-C{col+1}', projection_format)
