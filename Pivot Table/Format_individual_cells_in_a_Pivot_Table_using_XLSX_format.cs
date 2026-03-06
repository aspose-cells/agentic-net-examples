using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using System.Drawing;

namespace AsposeCellsPivotFormatting
{
    public class FormatIndividualPivotCell
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Product");
            sheet.Cells["C1"].PutValue("Sales");

            sheet.Cells["A2"].PutValue("Electronics");
            sheet.Cells["B2"].PutValue("Laptop");
            sheet.Cells["C2"].PutValue(1200);

            sheet.Cells["A3"].PutValue("Electronics");
            sheet.Cells["B3"].PutValue("Phone");
            sheet.Cells["C3"].PutValue(800);

            sheet.Cells["A4"].PutValue("Furniture");
            sheet.Cells["B4"].PutValue("Chair");
            sheet.Cells["C4"].PutValue(150);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Column, "Product");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Calculate the pivot data so that the layout is generated
            pivot.CalculateData();

            // Create a style to apply to a specific cell (e.g., the header cell for "Sales")
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.Name = "Arial";
            headerStyle.Font.Size = 12;
            headerStyle.Font.IsBold = true;
            headerStyle.ForegroundColor = Color.Yellow;
            headerStyle.Pattern = BackgroundType.Solid;

            // The header cell for the data field is located at row 1, column 1 in pivot coordinates
            // (row index is 1‑based in the pivot table view, column index is 0‑based)
            // Use the Format(row, column, Style) method to format that single cell
            pivot.Format(1, 1, headerStyle);

            // Save the workbook in XLSX format
            workbook.Save("FormattedPivotCell.xlsx", SaveFormat.Xlsx);
        }
    }
}