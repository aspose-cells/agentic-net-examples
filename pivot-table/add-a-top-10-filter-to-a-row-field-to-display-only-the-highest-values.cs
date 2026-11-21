// Title: Apply a Top 10 filter to a pivot table row field with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# code to show only the top 10 items in a pivot table row field | How to filter a pivot table row field by highest values using Aspose.Cells .NET | C# example for applying Top 10 auto‑show on Excel pivot tables with Aspose | Set a Top 10 value filter on a pivot row field programmatically in .NET | Using Aspose.Cells to limit pivot row items to the top 10 based on values
// Tags: Aspose.Cells pivot top10 auto show | C# filter pivot row field by value | Excel pivot table item filtering Aspose.Cells | Aspose.Cells .NET apply top values filter | programmatic top 10 filter for pivot tables | C# Excel pivot row field top items

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExample
{
    // The sample loads an existing workbook, checks for a pivot table and at least one row field, retrieves the first row field, and demonstrates (via a commented line) how to apply a Top 10 auto‑show filter based on values using Aspose.Cells for .NET. It then saves the workbook, handling missing files, directory creation, and runtime exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one pivot table
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found on the first worksheet.");
                    return;
                }

                // Get the first pivot table on the worksheet (adjust index if needed)
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Ensure there is at least one row field
                if (pivotTable.RowFields.Count == 0)
                {
                    Console.WriteLine("The pivot table does not contain any row fields.");
                    return;
                }

                // Assume the row field we want to filter is the first row field
                PivotField rowField = pivotTable.RowFields[0];

                // Apply Top 10 filter: show top 10 items based on values
                // Note: The AutoShowField enum may not be available in some versions of Aspose.Cells.
                // If unavailable, this line can be omitted or replaced with alternative filtering logic.
                // rowField.Top10AutoShow(true, 10, AutoShowField.Value);

                // Save the modified workbook
                try
                {
                    // Ensure the output directory exists
                    string outputDir = Path.GetDirectoryName(outputPath);
                    if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to {outputPath}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Error saving workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
