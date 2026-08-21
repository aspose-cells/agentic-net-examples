// Title: Remove a PivotTable and verify its deletion with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add a PivotTable, delete it using Worksheet.PivotTables.Remove, and confirm the worksheet contains zero pivot objects before saving the file.
// Keywords: Aspose.Cells delete pivot table | C# remove PivotTable | verify pivot table removal | Worksheet.PivotTables.Count | Aspose.Cells .NET pivot cleanup
// Common Searches: how to delete a pivot table in Aspose.Cells C# | check if worksheet has any pivot tables after removal | Aspose.Cells remove pivot and confirm count | C# code to clear pivot tables from a sheet
// Developer Intent: Delete an existing PivotTable and ensure the worksheet no longer contains any pivot objects.
// Use Cases: Clean up pivot tables after data updates to prevent stale references. | Validate workbook integrity before export by confirming zero pivot tables. | Implement conditional workflows that proceed only when a sheet is free of pivot objects.
// AI Prompts: Show C# code using Aspose.Cells to remove a specific PivotTable and assert that sheet.PivotTables.Count is zero. | Provide an example that iterates through all PivotTables in a worksheet, deletes each, and logs the result. | Explain safe removal of a PivotTable with Aspose.Cells, including handling cases where the table may not exist.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRemovalDemo
{
    // Demonstrates how to create a workbook, add a PivotTable, delete it using Worksheet.PivotTables.Remove, and confirm the worksheet contains zero pivot objects before saving the file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["A4"].PutValue("Apple");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(130);

            // Add a pivot table
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Product");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivot.RefreshData();
            pivot.CalculateData();

            // Remove the pivot table
            sheet.PivotTables.Remove(pivot);

            // Verify that no pivot tables remain in the worksheet
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("Pivot table removed successfully. No pivot tables remain.");
            }
            else
            {
                Console.WriteLine("Pivot table removal failed. Remaining count: " + sheet.PivotTables.Count);
            }

            // Save the workbook (optional, just to complete lifecycle)
            workbook.Save("PivotTableRemoved.xlsx");
        }
    }
}
