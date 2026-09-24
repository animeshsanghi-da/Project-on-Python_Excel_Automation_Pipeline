import xlwings as xw
import pandas as pd

def process_with_xlwings(excel_file):
    # Launch Excel application invisibly
    app = xw.App(visible=False)
    
    try:
        wb = app.books.open(excel_file)
        sales_sheet = wb.sheets["Sales_Data"]

        # 1. Read data from openpyxl output sheet into Pandas DataFrame
        # Evaluates existing table range dynamically
        data_range = sales_sheet.range("A1").expand("table")
        df = data_range.options(pd.DataFrame, header=1, index=False).value

        # 2. Perform Aggregation / Pivot logic using Pandas
        summary_df = df.groupby("Region").agg(
            Total_Units=("Units_Sold", "sum"),
            Total_Revenue=("Total_Revenue", "sum"),
            Total_Commission=("Commission", "sum")
        ).reset_index()

        # 3. Create or clear "Executive_Summary" Sheet
        if "Executive_Summary" in [sheet.name for sheet in wb.sheets]:
            summary_sheet = wb.sheets["Executive_Summary"]
            summary_sheet.clear()
        else:
            summary_sheet = wb.sheets.add("Executive_Summary", after=sales_sheet)

        # 4. Write Summary Header & Table (Merged A1:D1)
        title_range = summary_sheet.range("A1:D1")
        title_range.merge()  # Merges cells A1 through D1
        title_range.value = "Regional Executive Sales Summary"
        title_range.font.size = 14
        title_range.font.bold = True

        # Write DataFrame starting at cell A3
        summary_sheet.range("A3").options(index=False).value = summary_df

        # 5. Format Summary Table Live via xlwings
        header_range = summary_sheet.range("A3:D3")
        header_range.color = (31, 78, 120)  # Dark Navy Blue (RGB)
        header_range.font.color = (255, 255, 255)  # White text
        header_range.font.bold = True

        data_rows = len(summary_df)
        data_range_fmt = summary_sheet.range(f"C4:D{3 + data_rows}")
        data_range_fmt.number_format = "₹#,##0.00"

        # Auto-fit columns across the summary sheet
        summary_sheet.autofit()

        # Save changes
        wb.save()
        print(f"xlwings automation completed. 'Executive_Summary' tab updated in {excel_file}.")

    finally:
        # Cleanly quit Excel process
        wb.close()
        app.quit()

if __name__ == "__main__":
    # Runs xlwings on top of the file output by openpyxl
    process_with_xlwings("Cleaned_Sales_OpenPyXL.xlsx")