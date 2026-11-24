// Title: Refresh all pivot tables after loading an Excel workbook with ParsingPivotCachedRecords using Aspose.Cells for .NET
// Common Searches: c# aspnet cells load workbook with ParsingPivotCachedRecords and refresh pivot tables | how to enable parsing of pivot cached records when opening Excel with Aspose.Cells | programmatically refresh and recalculate all pivot tables after loading workbook in C# | Aspose.Cells LoadOptions ParsingPivotCachedRecords example for pivot table update | iterate worksheets and refresh pivot tables using Aspose.Cells for .NET | update pivot cache records after loading Excel file with Aspose.Cells C#
// Tags: ParsingPivotCachedRecords load option aspnet cells | refresh pivot tables aspnet cells c# | calculate pivot data aspnet cells | iterate worksheets pivot tables aspnet | update pivot cache after workbook load aspnet | aspnet cells pivot table RefreshData CalculateData

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The sample loads an Excel workbook with Aspose.Cells using LoadOptions where ParsingPivotCachedRecords is set to true, then iterates through each worksheet and its pivot tables, calling RefreshData() and CalculateData() to update and recalculate the pivot data, and finally saves the modified workbook.
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

            // Configure load options to parse pivot cached records
            LoadOptions loadOptions = new LoadOptions
            {
                ParsingPivotCachedRecords = true
            };

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Refresh all pivot tables in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pivotTable in sheet.PivotTables)
                {
                    // Refresh the data source of the pivot table
                    pivotTable.RefreshData();
                    // Recalculate the pivot table after refresh
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
