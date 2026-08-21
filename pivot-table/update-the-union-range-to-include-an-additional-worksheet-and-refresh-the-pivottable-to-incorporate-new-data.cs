// Title: Update PivotTable Union Data Source and Refresh It with Aspose.Cells for .NET (C#)
// Description: Loads an existing workbook, replaces the first PivotTable's source with a union range that spans Sheet1 and Sheet2, refreshes and recalculates the table, then saves the result to a new file.
// Keywords: Aspose.Cells PivotTable union range | change pivot data source C# | refresh pivot table Aspose.Cells | calculate pivot data .NET | multiple worksheet pivot source
// Common Searches: Aspose.Cells add second worksheet to pivot source | C# refresh PivotTable after changing data source | how to use union range for PivotTable in Aspose.Cells | update and recalculate PivotTable programmatically
// Developer Intent: Programmatically extend a PivotTable's source to include another worksheet and trigger a refresh so the new data is reflected.
// Use Cases: Merge sales figures from two sheets into a single PivotTable without manual re‑configuration. | Automate PivotTable updates in a reporting pipeline after new worksheet data is added. | Ensure calculations are current after modifying the PivotTable's underlying data range.
// AI Prompts: Write C# code using Aspose.Cells to change a PivotTable's data source to a union of Sheet1!A1:C10 and Sheet3!A1:C10, then refresh it. | Explain the steps required to refresh and recalculate a PivotTable after calling PivotTable.ChangeDataSource in Aspose.Cells. | Suggest robust error‑handling patterns when the additional worksheet referenced in a union range does not exist.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Loads an existing workbook, replaces the first PivotTable's source with a union range that spans Sheet1 and Sheet2, refreshes and recalculates the table, then saves the result to a new file.
    public class UpdateUnionRangeAndRefreshPivot
    {
        // Entry point for the application
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
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                throw new FileNotFoundException($"Input file not found: {inputPath}");
            }

            // Load the existing workbook that contains the pivot table
            Workbook workbook = new Workbook(inputPath);

            // Assume the pivot table is on the first worksheet
            Worksheet pivotWorksheet = workbook.Worksheets[0];

            // Ensure there is at least one pivot table
            if (pivotWorksheet.PivotTables.Count == 0)
            {
                throw new InvalidOperationException("No pivot tables found on the first worksheet.");
            }

            PivotTable pivotTable = pivotWorksheet.PivotTables[0];

            // Define the new union data source that includes an additional worksheet (e.g., "Sheet2")
            // The format is "SheetName!Range"
            string[] newDataSource = new string[]
            {
                "Sheet1!A1:C10",   // Existing range
                "Sheet2!A1:C10"    // Additional range to be included
            };

            // Change the pivot table's data source to the new union range
            pivotTable.ChangeDataSource(newDataSource);

            // Refresh the pivot table data (the method is not obsolete in current API)
            pivotTable.RefreshData();

            // Recalculate the pivot table after data refresh
            pivotTable.CalculateData();

            // Optionally, refresh all pivot tables in the worksheet
            // pivotWorksheet.RefreshPivotTables();

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
                throw;
            }
        }
    }
}
