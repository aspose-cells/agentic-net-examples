// Title: Refresh every PivotTable in an Excel workbook after bulk data changes using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, iterates through all worksheets, and updates each PivotTable's cache and recalculates its data. | Demonstrate how to handle missing input files and save the workbook after programmatically refreshing all PivotTables with Aspose.Cells.
// Common Searches: how to programmatically refresh all pivot tables in an Excel file using Aspose.Cells C# | Aspose.Cells loop through worksheets to update pivot caches after data import | C# example for RefreshData and CalculateData on multiple PivotTables in a workbook | save workbook after bulk pivot table refresh with Aspose.Cells for .NET | error handling when refreshing pivot tables in Aspose.Cells
// Tags: refresh pivot tables Aspose.Cells C# | iterate worksheets pivot tables Aspose.Cells | pivotcache RefreshData Aspose.Cells | calculate pivot data Aspose.Cells | save workbook after pivot refresh Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable and PivotCache classes

// The sample loads an existing Excel workbook, loops through each worksheet and each PivotTable, calls RefreshData and CalculateData to update the cache and recalculate values, then saves the modified file while handling missing files and runtime errors.
class RefreshPivotTables
{
    static void Main()
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

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each PivotTable in the worksheet
                foreach (PivotTable pivotTable in sheet.PivotTables)
                {
                    try
                    {
                        // Refresh the underlying PivotCache (correct API)
                        pivotTable.RefreshData();

                        // Recalculate the PivotTable values
                        pivotTable.CalculateData();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error refreshing pivot table '{pivotTable.Name}' on sheet '{sheet.Name}': {ex.Message}");
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
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
            // Handle any runtime exceptions
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
