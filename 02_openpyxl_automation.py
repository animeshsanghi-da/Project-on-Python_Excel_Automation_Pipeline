import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def process_with_openpyxl(input_file, output_file):
    # Load workbook and select active sheet
    wb = openpyxl.load_workbook(input_file)
    ws = wb["Sales_Data"]

    # 1. Add calculated column headers
    ws["G1"] = "Total_Revenue"
    ws["H1"] = "Commission"

    # Define Styles
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")
    
    # Define Summary Row Styles (Classic Accounting Style)
    summary_font = Font(name="Calibri", size=11, bold=True, color="000000")
    summary_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")  # Light Gray
    summary_border = Border(
        top=Side(style="thin", color="000000"),
        bottom=Side(style="double", color="000000")  # Accounting double underline
    )

    thin_border = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000")
    )

    max_row = ws.max_row

    # 2. Add Excel Formulas and Formatting to Data Rows
    for row in range(2, max_row + 1):
        # Insert native Excel formulas
        ws[f"G{row}"] = f"=E{row}*F{row}"  # Total_Revenue = Units_Sold * Unit_Price
        ws[f"H{row}"] = f"=IF(G{row}>2000, G{row}*0.08, G{row}*0.04)"  # 8% commission if > 2000, else 4%

        # Number Formatting (Currency)
        ws[f"F{row}"].number_format = "₹#,##0.00"
        ws[f"G{row}"].number_format = "₹#,##0.00"
        ws[f"H{row}"].number_format = "₹#,##0.00"

        # Cell Borders & Alignment
        for col in range(1, 9):
            cell = ws.cell(row=row, column=col)
            cell.border = thin_border
            if col in [1, 5]:  # ID and Units
                cell.alignment = center_align
            elif col in [6, 7, 8]:  # Financials
                cell.alignment = right_align

    # 3. Add Summary Row at the Bottom
    summary_row = max_row + 1
    ws[f"A{summary_row}"] = "Total"
    ws[f"G{summary_row}"] = f"=SUM(G2:G{max_row})"
    ws[f"H{summary_row}"] = f"=SUM(H2:H{max_row})"

    # Style Summary Row
    total_font = Font(name="Calibri", size=11, bold=True)
    ws[f"A{summary_row}"].font = total_font
    ws[f"G{summary_row}"].font = total_font
    ws[f"H{summary_row}"].font = total_font
    ws[f"G{summary_row}"].number_format = "₹#,##0.00"
    ws[f"H{summary_row}"].number_format = "₹#,##0.00"

    # 4. Style Headers
    for col in range(1, 9):
        cell = ws.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align

    # 5. Style Summary Row
    for col in range(1, 9):
        cell = ws.cell(row=summary_row, column=col)
        cell.font = summary_font
        cell.fill = summary_fill
        cell.border = summary_border

    # 6. Auto-fit Column Widths
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    # Save output
    wb.save(output_file)
    print(f"OpenPyXL automation completed. Saved as '{output_file}'")

if __name__ == "__main__":
    process_with_openpyxl("Raw_Sales_Data.xlsx", "Cleaned_Sales_OpenPyXL.xlsx")