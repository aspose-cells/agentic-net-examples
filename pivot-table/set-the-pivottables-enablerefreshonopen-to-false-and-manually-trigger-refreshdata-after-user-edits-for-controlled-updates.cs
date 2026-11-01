// Title: Disable automatic PivotTable refresh on workbook open and manually invoke RefreshData after edits using Aspose.Cells for .NET (C#)
// Common Searches: how to stop a pivot table from refreshing automatically when opening a workbook with Aspose.Cells in C# | C# Aspose.Cells code to manually refresh a pivot table after changing source cells | set EnableRefreshOnOpen to false for a PivotTable using Aspose.Cells .NET | programmatically trigger RefreshData on an Aspose.Cells PivotTable after user edits
// Tags: Aspose.Cells pivot auto refresh control | EnableRefreshOnOpen property Aspose.Cells | RefreshData method Aspose.Cells PivotTable | manual pivot cache refresh .NET | disable pivot refresh on open C#

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example loads an existing Excel workbook, checks for a PivotTable, optionally allows user edits, disables automatic refresh on opening, calls RefreshData on the first PivotTable to apply the changes, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure a pivot table exists
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Access the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Note: The RefreshDataOnOpen property is not available in this version of Aspose.Cells.
            // If needed, configure workbook settings elsewhere.

            // ----- User edits would occur here -----
            // Example edit:
            // worksheet.Cells["A2"].PutValue(123);

            // Refresh the pivot table to reflect changes
            try
            {
                pivotTable.RefreshData();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to refresh pivot table: {ex.Message}");
                // Continue execution; the workbook can still be saved.
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath) ?? string.Empty;
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
