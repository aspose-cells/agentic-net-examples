// Title: Load an Excel workbook with ParsingPivotCachedRecords enabled and refresh all pivot tables using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# load workbook with ParsingPivotCachedRecords and refresh pivot tables | how to enable ParsingPivotCachedRecords in LoadOptions and update pivot caches in .NET | refresh all pivot tables after loading an .xlsx file with Aspose.Cells | C# iterate worksheets and call RefreshData on each pivot table using Aspose.Cells | save workbook after refreshing pivot tables with Aspose.Cells LoadOptions
// Tags: ParsingPivotCachedRecords load option Aspose.Cells | pivot table refresh Aspose.Cells C# | iterate worksheets pivot tables Aspose.Cells | save workbook after pivot refresh Aspose.Cells | load Excel with pivot cache records Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example demonstrates loading an Excel file with LoadOptions.ParsingPivotCachedRecords set to true, iterating through each worksheet to refresh every PivotTable via RefreshData, and then saving the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Configure load options to parse pivot cached records
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                ParsingPivotCachedRecords = true
            };

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Refresh all pivot tables in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pivot in sheet.PivotTables)
                {
                    try
                    {
                        // Refresh the data source and recalculate the pivot table
                        pivot.RefreshData();
                        // In newer versions RefreshData() updates the cache; additional refresh is not required.
                    }
                    catch (Exception exPivot)
                    {
                        Console.WriteLine($"Failed to refresh pivot table '{pivot.Name}' on sheet '{sheet.Name}': {exPivot.Message}");
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the refreshed workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
