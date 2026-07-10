using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].Value = "Category";
        sheet.Cells["B1"].Value = "Amount";
        sheet.Cells["A2"].Value = "Fruit";
        sheet.Cells["B2"].Value = 100;
        sheet.Cells["A3"].Value = "Vegetable";
        sheet.Cells["B3"].Value = 150;
        sheet.Cells["A4"].Value = "Fruit";
        sheet.Cells["B4"].Value = 200;
        sheet.Cells["A5"].Value = "Vegetable";
        sheet.Cells["B5"].Value = 250;

        // Add a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add fields to the pivot table (row and data)
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

        // Restrict filter selections: disallow multiple filters per field
        pivotTable.AllowMultipleFiltersPerField = false;

        // Save the workbook to a file
        workbook.Save("PivotTable_AllowMultipleFiltersFalse.xlsx");
    }
}