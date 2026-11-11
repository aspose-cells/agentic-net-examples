// Title: Refresh all nested pivot tables in an Excel workbook using Aspose.Cells for .NET
// Common Searches: C# code to refresh every pivot table in an Excel file with Aspose.Cells | How to update child pivot tables after changing source data using Aspose.Cells .NET | Programmatically iterate and refresh nested pivot tables in a workbook with Aspose.Cells | Aspose.Cells .NET refresh all pivot tables before saving the workbook | Recursive refresh of pivot tables in Excel using the Aspose.Cells API
// Tags: Aspose.Cells .NET refresh all pivot tables | recursive pivot table refresh in Excel | C# iterate worksheets pivot tables Aspose.Cells | update nested pivot tables programmatically | refresh pivot data source Aspose.Cells API

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable class

namespace PivotTableRefreshExample
{
    // The example loads an Excel workbook, walks through each worksheet and its PivotTable collection, calls RefreshData on every pivot table to synchronize with the source data, and then saves the workbook, handling missing files and runtime errors.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Determine input and output file paths
                string inputPath = args.Length > 0 ? args[0] : "Input.xlsx";
                string outputPath = args.Length > 1 ? args[1] : "Output.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    throw new FileNotFoundException($"Input file not found: {inputPath}");
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Refresh all pivot tables in the workbook
                RefreshAllPivotTables(workbook);

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Log the exception details
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        /// <param name="workbook">The Aspose.Cells workbook instance.</param>
        private static void RefreshAllPivotTables(Workbook workbook)
        {
            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each pivot table on the current worksheet
                foreach (PivotTable pivotTable in sheet.PivotTables)
                {
                    try
                    {
                        // Refresh the pivot table's data source
                        pivotTable.RefreshData();
                    }
                    catch (Exception ex)
                    {
                        // Log any pivot-specific errors but continue processing others
                        Console.Error.WriteLine($"Failed to refresh pivot table '{pivotTable.Name}' on sheet '{sheet.Name}': {ex.Message}");
                    }
                }
            }
        }
    }
}
