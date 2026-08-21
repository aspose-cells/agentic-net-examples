// Title: C# – Refresh All Pivot Tables in an Aspose.Cells Workbook After Updating Source Data
// Description: Demonstrates how to modify worksheet data, then call Workbook.Worksheets.RefreshPivotTables() to recalculate every pivot table before saving the file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells refresh pivot tables C# | RefreshPivotTables method .NET | update pivot table after data change | programmatic pivot table refresh Aspose | refresh all pivots workbook | Aspose.Cells pivot table example
// Common Searches: how to refresh all pivot tables in Aspose.Cells C# | Aspose.Cells RefreshPivotTables usage | refresh pivot tables after editing source data .NET | programmatically update pivot tables Aspose.Cells | C# code to recalculate all pivots in workbook
// Developer Intent: Recalculate every pivot table in the workbook so it reflects the latest values in the underlying worksheet cells.
// Use Cases: After bulk updating sales figures, invoke RefreshPivotTables to keep all report pivots accurate before export. | When generating a multi‑sheet financial report that contains several pivot tables, use a single call to refresh them all in one step. | In an automated ETL pipeline that writes new data into a worksheet, call RefreshPivotTables to synchronize all analytical views.
// AI Prompts: Generate C# code that changes cell values and then refreshes all pivot tables using Aspose.Cells. | Show the difference between refreshing a single pivot table and refreshing all pivot tables in an Aspose.Cells workbook. | Provide best‑practice error handling for Workbook.Worksheets.RefreshPivotTables when a workbook may contain no pivots.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates how to modify worksheet data, then call Workbook.Worksheets.RefreshPivotTables() to recalculate every pivot table before saving the file with Aspose.Cells for .NET.
    public class RefreshAllPivotTablesDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate source data for the pivot table
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(1000);
                sheet.Cells["A3"].PutValue("Orange");
                sheet.Cells["B3"].PutValue(2000);
                sheet.Cells["A4"].PutValue("Banana");
                sheet.Cells["B4"].PutValue(3000);

                // Add a pivot table that uses the source data
                int pivotIndex = sheet.PivotTables.Add("A1:B4", "E3", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product as row field
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Sales as data field

                // Modify the underlying data after the pivot table has been created
                sheet.Cells["B2"].PutValue(1500); // Updated sales for Apple
                sheet.Cells["B3"].PutValue(2500); // Updated sales for Orange

                // Refresh all pivot tables in the workbook to reflect the data changes
                workbook.Worksheets.RefreshPivotTables();

                // Save the updated workbook
                workbook.Save("RefreshedPivotTables.xlsx");
                Console.WriteLine("Workbook saved as RefreshedPivotTables.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            RefreshAllPivotTablesDemo.Run();
        }
    }
}
