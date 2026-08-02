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
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(200);
        sheet.Cells["A4"].PutValue("Apple");
        sheet.Cells["B4"].PutValue(150);

        // Add a pivot table to the worksheet
        int ptIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[ptIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Delete the pivot table
        sheet.PivotTables.Remove(pivotTable);

        // Verify that the worksheet no longer contains any pivot tables
        if (sheet.PivotTables.Count == 0)
        {
            Console.WriteLine("Pivot table removed successfully. No pivot tables remain.");
        }
        else
        {
            Console.WriteLine("Pivot table removal failed. Remaining count: " + sheet.PivotTables.Count);
        }

        // Save the workbook
        workbook.Save("PivotTableRemoved.xlsx");
    }
}