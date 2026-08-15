// Title: Disable Auto‑Refresh for Aspose.Cells PivotTable and Refresh Manually in C#
// Description: Demonstrates how to set a PivotTable's RefreshDataOnOpeningFile to false, edit source data, then call RefreshData and CalculateData to update the pivot before saving the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells PivotTable manual refresh | RefreshDataOnOpeningFile false | C# Aspose.Cells pivot cache | disable automatic pivot refresh | RefreshData after source edit
// Common Searches: Aspose.Cells prevent pivot table auto refresh | how to manually refresh pivot table in C# | RefreshDataOnOpeningFile property usage | control pivot refresh Aspose.Cells .NET | update pivot after editing source cells
// Developer Intent: Turn off automatic pivot refresh on workbook open and trigger a controlled refresh after source data changes.
// Use Cases: Web apps where users modify data and the pivot should refresh only after submission, reducing latency. | Batch processing that updates many rows before recalculating the pivot to ensure accurate totals. | Generating reports that require a stable snapshot of source data until all calculations are finalized.
// AI Prompts: Write C# code with Aspose.Cells to create a pivot table, disable auto‑refresh, modify a cell, and manually refresh the pivot. | Explain the relationship between RefreshDataOnOpeningFile and RefreshData in Aspose.Cells and suggest best practices for performance. | Provide error‑handling patterns for saving a workbook after a manual pivot refresh in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates how to set a PivotTable's RefreshDataOnOpeningFile to false, edit source data, then call RefreshData and CalculateData to update the pivot before saving the workbook using Aspose.Cells for .NET.
    public class PivotTableControlledRefreshDemo
    {
        public static void Main(string[] args)
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
            Worksheet sheet = workbook.Worksheets[0];

            // Populate source data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("A");
            sheet.Cells["B4"].PutValue(30);

            // Add a pivot table based on the source range
            int ptIndex = sheet.PivotTables.Add("A1:B4", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[ptIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

            // Disable automatic refresh when the workbook is opened
            pivotTable.RefreshDataOnOpeningFile = false;

            // Initial calculation to populate the pivot table
            pivotTable.CalculateData();

            // Simulate a user editing the source data
            sheet.Cells["B2"].PutValue(100); // Change value from 10 to 100

            // Manually refresh the pivot cache and recalculate the pivot table
            pivotTable.RefreshData();   // Refreshes data from the source
            pivotTable.CalculateData(); // Recalculates the pivot based on refreshed data

            // Save the workbook
            string outputPath = "PivotTableControlledRefreshDemo.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
