using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the pivot table
        sheet.Cells["A1"].PutValue("Fruit");
        sheet.Cells["B1"].PutValue("Quantity");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["B3"].PutValue(15);
        sheet.Cells["A4"].PutValue("Banana");
        sheet.Cells["B4"].PutValue(20);

        // Create and configure the pivot table
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "FruitPivot");
        PivotTable pivot = sheet.PivotTables[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
        pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");

        // Save the workbook to XLSX format using default save options
        workbook.Save("PivotTableDemo.xlsx", SaveFormat.Xlsx);
    }
}