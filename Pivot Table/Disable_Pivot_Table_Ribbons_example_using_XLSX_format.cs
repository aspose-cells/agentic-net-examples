using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class DisablePivotTableRibbons
{
    public static void Main()
    {
        Run();
    }

    public static void Run()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(1000);
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["B3"].PutValue(1500);
        sheet.Cells["A4"].PutValue("Banana");
        sheet.Cells["B4"].PutValue(2000);

        // Add a pivot table to the worksheet
        int ptIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[ptIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product as row field
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Sales as data field

        // Save the workbook in XLSX format
        workbook.Save("DisablePivotTableRibbons.xlsx", SaveFormat.Xlsx);
    }
}