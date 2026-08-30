// Title: Disable automatic PivotTable refresh on workbook open and trigger manual refresh after data edits using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a PivotTable, sets RefreshDataOnOpeningFile to false, enables ManualUpdate, modifies source cells, and then calls RefreshData and CalculateData with Aspose.Cells. | Show a complete Aspose.Cells example that disables auto‑refresh for a PivotTable, updates the underlying data, and performs a manual refresh to recalculate the pivot.
// Common Searches: Aspose.Cells C# disable pivot table auto refresh on file open | Manually refresh Aspose.Cells PivotTable after editing source data | Set PivotTable.ManualUpdate true in Aspose.Cells .NET example | How to use RefreshDataOnOpeningFile property with Aspose.Cells PivotTable | Control when a PivotTable recalculates using Aspose.Cells
// Tags: Aspose.Cells pivot manual update | turn off pivot auto refresh Aspose.Cells | use RefreshDataOnOpeningFile property | invoke PivotTable.CalculateData programmatically | controlled pivot cache refresh .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The example demonstrates creating a workbook, adding source data, inserting a PivotTable, disabling automatic refresh on opening by setting RefreshDataOnOpeningFile to false, enabling ManualUpdate, performing an initial calculation, modifying source cells, and then manually refreshing the pivot cache with RefreshData followed by recalculating the pivot with CalculateData before saving the file.
    public class ControlledPivotRefreshDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];

            // Populate source data for the pivot table
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Amount");
            dataSheet.Cells["A2"].PutValue("A");
            dataSheet.Cells["B2"].PutValue(100);
            dataSheet.Cells["A3"].PutValue("B");
            dataSheet.Cells["B3"].PutValue(200);
            dataSheet.Cells["A4"].PutValue("A");
            dataSheet.Cells["B4"].PutValue(150);
            dataSheet.Cells["A5"].PutValue("B");
            dataSheet.Cells["B5"].PutValue(250);

            // Add a pivot table on the same sheet
            int pivotIndex = dataSheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
            PivotTable pivotTable = dataSheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category column
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Amount column

            // Disable automatic refresh when the workbook is opened
            pivotTable.RefreshDataOnOpeningFile = false;

            // Enable manual update so the pivot does not recalculate automatically after data changes
            pivotTable.ManualUpdate = true;

            // Initial calculation to populate the pivot table
            pivotTable.CalculateData();

            // Simulate user editing the source data
            dataSheet.Cells["B2"].PutValue(120); // Change amount for Category A
            dataSheet.Cells["B3"].PutValue(220); // Change amount for Category B

            // Manually refresh the pivot cache and recalculate the pivot table
            pivotTable.RefreshData();   // Refreshes the cache from the data source
            pivotTable.CalculateData(); // Recalculates the pivot based on the refreshed cache

            // Save the workbook
            string outputPath = "ControlledPivotRefreshDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
