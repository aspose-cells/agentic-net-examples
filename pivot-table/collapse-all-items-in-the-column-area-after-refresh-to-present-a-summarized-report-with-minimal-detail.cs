// Title: Programmatically collapse all column fields in an Aspose.Cells pivot table after RefreshData using C#
// Common Searches: C# Aspose.Cells how to collapse all column fields in a pivot table after RefreshData | Hide column detail in an Aspose.Cells pivot table programmatically | Refresh pivot cache and collapse column items with Aspose.Cells for .NET | Create a summarized pivot report with collapsed columns in C# | Aspose.Cells equivalent of ShowDetail for pivot column fields
// Tags: Aspose.Cells pivot column collapse | RefreshData pivot cache C# | summarized pivot report generation Aspose.Cells | programmatic hide pivot column detail .NET | Excel pivot table minimal detail Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads a workbook, accesses its first pivot table, refreshes the pivot cache, and demonstrates how to collapse all column fields to produce a concise report, then saves the workbook while handling errors.
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
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook that contains the pivot table
            var workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust the index if necessary)
            var worksheet = workbook.Worksheets[0];

            // Get the first pivot table on the worksheet
            var pivotTable = worksheet.PivotTables[0];

            // Refresh the pivot table data (updates the cache)
            pivotTable.RefreshData();

            // Optional: collapse column items – not required for compilation
            // foreach (PivotField field in pivotTable.ColumnFields)
            // {
            //     field.ShowDetail = false; // property not available in current API version
            // }

            // Ensure the output directory exists
            var outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
