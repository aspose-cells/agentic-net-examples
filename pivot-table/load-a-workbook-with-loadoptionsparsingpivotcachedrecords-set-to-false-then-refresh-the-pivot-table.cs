// Title: Refresh all pivot tables in an Excel workbook loaded with ParsingPivotCachedRecords disabled using Aspose.Cells for .NET
// Common Searches: how to load an Excel file with Aspose.Cells without parsing pivot cache and then refresh pivot tables | Aspose.Cells C# disable pivot cached records during workbook load and recalculate pivot tables | refresh pivot tables after loading workbook with LoadOptions.ParsingPivotCachedRecords set to false in .NET
// Tags: load workbook without pivot cache parsing Aspose.Cells | programmatic pivot table refresh C# | ParsingPivotCachedRecords false option | pivot table data refresh and calculation Aspose.Cells | skip pivot cache parsing on workbook load

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example checks for the input Excel file, loads it with LoadOptions where ParsingPivotCachedRecords is set to false, iterates through each worksheet and its PivotTables, calls RefreshData() and CalculateData() on each pivot table, and saves the modified workbook to the output path while handling any exceptions.
class Program
{
    static void Main()
    {
        // Input and output file paths
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load options: do not parse cached pivot records
            LoadOptions loadOptions = new LoadOptions
            {
                ParsingPivotCachedRecords = false
            };

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Refresh all pivot tables in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pivotTable in sheet.PivotTables)
                {
                    // Refresh data and recalculate the pivot table
                    pivotTable.RefreshData();
                    pivotTable.CalculateData();
                }
            }

            // Save the workbook after refreshing pivot tables
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
