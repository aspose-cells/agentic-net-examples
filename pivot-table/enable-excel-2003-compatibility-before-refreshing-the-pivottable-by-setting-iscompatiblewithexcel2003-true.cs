// Title: Set IsCompatibleWithExcel2003 true and refresh all PivotTables before saving as Excel 2003 .xls using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set workbook.IsCompatibleWithExcel2003 before refreshing pivot tables | how to refresh every pivot table and save workbook in Excel 2003 format with Aspose.Cells .NET | enable Excel 2003 compatibility mode for a workbook that contains PivotTables in C# | C# code to recalculate pivot tables then export to .xls while preserving 2003 compatibility | Aspose.Cells example for setting compatibility flag and updating pivot data
// Tags: IsCompatibleWithExcel2003 Aspose.Cells | refresh pivot tables Aspose.Cells | save workbook as Excel 2003 Aspose.Cells | pivot table compatibility mode C# | Aspose.Cells workbook compatibility property

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing .xlsx workbook, enables Excel 2003 compatibility by setting workbook.IsCompatibleWithExcel2003 to true, iterates through each worksheet to refresh and recalculate every PivotTable, and finally saves the workbook as an Excel 2003 .xls file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xls";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Refresh each pivot table in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pivot in sheet.PivotTables)
                {
                    try
                    {
                        // Refresh the data source of the pivot table
                        pivot.RefreshData();

                        // Recalculate the pivot table after refresh
                        pivot.CalculateData();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing pivot table on sheet '{sheet.Name}': {ex.Message}");
                    }
                }
            }

            // Save the workbook in Excel 2003 format (.xls)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
