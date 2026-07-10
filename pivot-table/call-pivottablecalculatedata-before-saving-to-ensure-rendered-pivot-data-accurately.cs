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
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Amount");
        sheet.Cells["A2"].PutValue("Fruit");
        sheet.Cells["B2"].PutValue(50);
        sheet.Cells["A3"].PutValue("Vegetable");
        sheet.Cells["B3"].PutValue(30);
        sheet.Cells["A4"].PutValue("Fruit");
        sheet.Cells["B4"].PutValue(70);
        sheet.Cells["A5"].PutValue("Vegetable");
        sheet.Cells["B5"].PutValue(40);

        // Add a pivot table to the worksheet
        int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Configure the pivot table fields
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

        // Calculate pivot data before saving to ensure correct rendered values
        pivotTable.CalculateData();

        // Save the workbook
        workbook.Save("PivotCalculated.xlsx");
    }
}