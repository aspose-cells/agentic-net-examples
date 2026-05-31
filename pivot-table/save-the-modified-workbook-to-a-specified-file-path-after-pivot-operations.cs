using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class SavePivotWorkbookDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data for the pivot table
        cells["A1"].Value = "Category";
        cells["B1"].Value = "Amount";
        cells["A2"].Value = "Food";
        cells["B2"].Value = 100;
        cells["A3"].Value = "Drink";
        cells["B3"].Value = 150;
        cells["A4"].Value = "Supplies";
        cells["B4"].Value = 200;

        // Add a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "SalesPivot");
        PivotTable pivot = sheet.PivotTables[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, "Category");
        pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

        // Refresh the pivot table to ensure it reflects the source data
        sheet.RefreshPivotTables();

        // Save the modified workbook to the specified file path
        workbook.Save("ModifiedPivotWorkbook.xlsx");
    }
}