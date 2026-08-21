// Title: C# – Delete a Pivot Table and Verify No Remaining Pivot Objects with Aspose.Cells for .NET
// Description: This Aspose.Cells for .NET example creates a workbook, adds a pivot table on sample data, removes it using PivotTables.Remove (and optionally RemoveAt), checks that the worksheet’s PivotTables count is zero, and saves the file without any pivot objects.
// Keywords: Aspose.Cells delete pivot table C# | remove pivot table .NET | PivotTables.Remove Aspose.Cells | PivotTables.RemoveAt example | verify zero pivot tables | C# Aspose.Cells pivot removal | Aspose.Cells workbook without pivots
// Common Searches: how to delete a pivot table using Aspose.Cells C# | check if worksheet has pivot tables after removal | Aspose.Cells PivotTables.Remove vs RemoveAt | remove all pivot tables from a workbook Aspose | sample code for deleting pivot tables in .NET
// Developer Intent: Remove an existing pivot table from a worksheet and confirm that the worksheet contains no pivot tables.
// Use Cases: Erase a temporary pivot table before exporting the workbook to PDF or CSV. | Clear outdated pivot tables when regenerating a report with fresh data. | Programmatically ensure a workbook is free of pivot objects before sending it to a client that does not support pivots.
// AI Prompts: Write C# code that deletes a pivot table by name with Aspose.Cells and returns the remaining pivot count. | Explain when to use PivotTables.Remove versus PivotTables.RemoveAt in Aspose.Cells. | Create a reusable method that removes every pivot table from a given worksheet and logs the before/after counts.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // This Aspose.Cells for .NET example creates a workbook, adds a pivot table on sample data, removes it using PivotTables.Remove (and optionally RemoveAt), checks that the worksheet’s PivotTables count is zero, and saves the file without any pivot objects.
    public class DeletePivotTableDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
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

            // Add a pivot table to the worksheet
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh and calculate data using the correct API
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Verify that the pivot table exists
            Console.WriteLine("Pivot tables count before removal: " + sheet.PivotTables.Count);

            // Remove the pivot table using the Remove method (deletes data as well)
            sheet.PivotTables.Remove(pivotTable);

            // Confirm that no pivot tables remain in the worksheet
            Console.WriteLine("Pivot tables count after removal: " + sheet.PivotTables.Count);

            // (Optional) Demonstrate removal by index if another pivot table is added
            // Add a second pivot table
            int secondIndex = sheet.PivotTables.Add("A1:B4", "D10", "SecondPivot");
            // Remove it using RemoveAt (index 0 because it's the only one now)
            sheet.PivotTables.RemoveAt(0);
            Console.WriteLine("Pivot tables count after second removal: " + sheet.PivotTables.Count);

            // Save the workbook (the file will contain no pivot tables)
            workbook.Save("PivotTableDeleted.xlsx");
        }
    }
}
