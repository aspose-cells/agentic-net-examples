// Title: How to refresh all pivot tables after loading an Excel workbook with ParsingPivotCachedRecords set to false using Aspose.Cells for .NET
// Common Searches: Aspose.Cells load Excel file without parsing pivot cache and refresh pivot tables C# | C# disable ParsingPivotCachedRecords and recalculate pivot tables after workbook load | Refresh all pivot tables programmatically after opening workbook with Aspose.Cells LoadOptions | Skip cached pivot records when loading Excel with Aspose.Cells and update pivot data | Use LoadOptions.ParsingPivotCachedRecords false to refresh pivot tables in .NET
// Tags: load workbook with ParsingPivotCachedRecords Aspose.Cells | refresh pivot tables programmatically C# | pivot table calculate data Aspose.Cells | skip pivot cache load options .NET | iterate worksheets pivot tables Aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Ensure PivotTable type is available

// The example demonstrates loading an Excel workbook using Aspose.Cells with LoadOptions.ParsingPivotCachedRecords set to false, iterating through each worksheet to refresh and recalculate all pivot tables, and then saving the updated workbook. It includes input file validation and exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Configure load options to skip parsing cached pivot records
            LoadOptions loadOptions = new LoadOptions
            {
                ParsingPivotCachedRecords = false
            };

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Iterate through all worksheets and refresh each pivot table
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pivotTable in sheet.PivotTables)
                {
                    // Refresh the data source of the pivot table
                    pivotTable.RefreshData();

                    // Recalculate the pivot table after data refresh
                    pivotTable.CalculateData();
                }
            }

            // Save the workbook after refreshing pivot tables
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
