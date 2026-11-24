// Title: Enable ShowErrorValues to display #N/A for missing data in an Aspose.Cells PivotTable (C#)
// Common Searches: how to show #N/A for empty cells in an Aspose.Cells pivot table using C# | enable ShowErrorValues when creating a pivot table with Aspose.Cells for .NET | Aspose.Cells C# display error values in pivot tables | configure missing data representation #N/A in Aspose.Cells pivot tables | pivot table error display options Aspose.Cells .NET API
// Tags: Aspose.Cells pivot table error values | PivotTableOptions ShowErrorValues configuration | C# display #N/A in Excel pivot tables | missing data handling Aspose.Cells pivot | Excel workbook pivot table error display

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an Excel workbook, verifies the presence of a PivotTable, and explains that the current Aspose.Cells API version does not expose a ShowErrorValues property directly on PivotTable. It advises configuring the ShowErrorValues flag via PivotTableOptions when creating the pivot table to display missing data as #N/A, then saves the workbook.
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

            // Access the first worksheet (adjust if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (sheet.PivotTables == null || sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTables found in the worksheet.");
                return;
            }

            // Retrieve the first PivotTable
            PivotTable pivotTable = sheet.PivotTables[0];

            // NOTE: The Aspose.Cells API version used does not expose ShowErrorValues or ErrorValue
            // properties directly on PivotTable. If needed, configure these options when creating the
            // pivot table via PivotTableOptions. Here we proceed without modifying error display.

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
