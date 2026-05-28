using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class DisablePivotAutoFitDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Amount");
        sheet.Cells["A2"].PutValue("Food");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["A3"].PutValue("Transport");
        sheet.Cells["B3"].PutValue(80);
        sheet.Cells["A4"].PutValue("Utilities");
        sheet.Cells["B4"].PutValue(150);

        // Add a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "MyPivot");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Configure the pivot fields
        pivot.AddFieldToArea(PivotFieldType.Row, "Category");
        pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
        pivot.CalculateData();

        // Manually set column widths for the pivot table area (optional)
        sheet.Cells.SetColumnWidth(3, 20); // Column D
        sheet.Cells.SetColumnWidth(4, 25); // Column E

        // Disable automatic column width fitting when the pivot table is refreshed or updated
        pivot.AutofitColumnWidthOnUpdate = false;

        // Refresh pivot tables to ensure the setting takes effect
        sheet.RefreshPivotTables();

        // Save the workbook with the disabled autofit setting
        workbook.Save("PivotAutoFitDisabled.xlsx");
    }
}