// Title: How to disable PivotTable auto‑refresh on workbook open and manually refresh data after source edits using Aspose.Cells for .NET
// Common Searches: Aspose.Cells set EnableRefreshOnOpen false for PivotTable in C# | Refresh PivotTable programmatically after modifying source cells with Aspose.Cells | Prevent PivotTable from refreshing automatically when opening workbook using Aspose.Cells | C# code to turn off PivotTable auto refresh and call RefreshData manually
// Tags: Aspose.Cells PivotTable DisableAutoRefresh | EnableRefreshOnOpen property usage | PivotTable RefreshData method .NET | Manual pivot data refresh Aspose.Cells | Update source data and refresh PivotTable C#

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example loads a workbook, disables the PivotTable's EnableRefreshOnOpen flag, edits source data, calls RefreshData to apply the changes, and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook containing the PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Access the worksheet that holds the PivotTable (adjust name if necessary)
            Worksheet pivotSheet = workbook.Worksheets["PivotSheet"];
            if (pivotSheet == null)
            {
                Console.WriteLine("Error: Worksheet \"PivotSheet\" not found.");
                return;
            }

            // Retrieve the first PivotTable in the worksheet
            if (pivotSheet.PivotTables.Count == 0)
            {
                Console.WriteLine("Error: No PivotTables found in \"PivotSheet\".");
                return;
            }
            PivotTable pivotTable = pivotSheet.PivotTables[0];

            // ----- Simulate user editing the source data -----
            Worksheet sourceSheet = workbook.Worksheets["SourceData"];
            if (sourceSheet == null)
            {
                Console.WriteLine("Error: Worksheet \"SourceData\" not found.");
                return;
            }
            // Example edit: change the value in cell B2
            sourceSheet.Cells["B2"].PutValue(12345);
            // ------------------------------------------------

            // Refresh the PivotTable to reflect the edited data
            try
            {
                pivotTable.RefreshData();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Unable to refresh PivotTable data. {ex.Message}");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
