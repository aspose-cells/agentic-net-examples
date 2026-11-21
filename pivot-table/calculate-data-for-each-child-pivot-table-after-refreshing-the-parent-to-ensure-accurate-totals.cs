// Title: Refresh parent pivot table and recalculate all child pivot tables in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells refresh parent pivot table and update child pivots | how to programmatically recalculate all pivot tables after refreshing with Aspose.Cells | Aspose.Cells .NET iterate worksheets to refresh and calculate pivot tables | use RefreshData and CalculateData for multiple pivot tables in Aspose.Cells C# | update Excel workbook pivot tables after data change with Aspose.Cells
// Tags: Aspose.Cells refresh pivot table .NET | calculate pivot data Aspose.Cells C# | iterate worksheets pivot tables Aspose.Cells | parent child pivot table synchronization Aspose.Cells | save workbook after pivot refresh Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable class

namespace AsposeCellsExamples
{
    // The example loads an Excel workbook, refreshes and calculates the first pivot table (treated as the parent), then iterates through all worksheets to refresh and calculate every other pivot table (children), and finally saves the updated workbook to the specified output path.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths (provide via command‑line args or replace with actual paths)
            string inputPath = args.Length > 0 ? args[0] : "input.xlsx";
            string outputPath = args.Length > 1 ? args[1] : "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Assume the parent pivot table is the first one in the first worksheet
                Worksheet parentSheet = workbook.Worksheets[0];
                if (parentSheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found in the first worksheet.");
                    return;
                }

                PivotTable parentPivot = parentSheet.PivotTables[0];

                // Refresh and calculate the parent pivot table
                parentPivot.RefreshData();
                parentPivot.CalculateData();

                // Refresh and calculate all other pivot tables
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (PivotTable pivot in sheet.PivotTables)
                    {
                        if (pivot == parentPivot)
                            continue;

                        pivot.RefreshData();
                        pivot.CalculateData();
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the updated workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
