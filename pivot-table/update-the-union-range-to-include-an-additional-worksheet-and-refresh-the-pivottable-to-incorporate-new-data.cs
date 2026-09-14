// Title: Add a worksheet to a PivotTable union data source and refresh the PivotTable with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that adds Sheet3 to an existing PivotTable's union data source and then refreshes the PivotTable using Aspose.Cells. | Demonstrate how to call ChangeDataSource with multiple sheet ranges and invoke RefreshPivotTables in Aspose.Cells for .NET. | Show the steps to load a workbook, modify a PivotTable's union source to include an extra worksheet, refresh it, and save the file with Aspose.Cells.
// Common Searches: asp.net aspose.cells add sheet to pivot table union source | c# refresh pivot table after updating union data source aspose.cells | how to include multiple worksheets in a pivot table source using Aspose.Cells | example of ChangeDataSource with union ranges in Aspose.Cells .NET | update pivot table source range to include Sheet3 and refresh in C#
// Tags: pivot table change union data source aspose.cells | add worksheet to pivot table source c# | refresh pivot tables after data source modification aspose.cells | changeDataSource multi-sheet range aspose.cells | c# update pivot table union range

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Loads an existing workbook, expands the PivotTable's union data source to include Sheet3 by calling ChangeDataSource with a multi-sheet range, refreshes the PivotTable, and saves the updated file.
    public class UpdateUnionRangeAndRefreshPivot
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "OutputWorkbook.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");
            }

            // Load the existing workbook containing a PivotTable with a union data source
            Workbook workbook = new Workbook(inputPath);

            // Assume the PivotTable is on the first worksheet
            Worksheet pivotWorksheet = workbook.Worksheets[0];

            // Get the first PivotTable in that worksheet
            if (pivotWorksheet.PivotTables.Count == 0)
            {
                throw new InvalidOperationException("No PivotTable found on the first worksheet.");
            }
            PivotTable pivotTable = pivotWorksheet.PivotTables[0];

            // Define the new union range that includes an additional worksheet (Sheet3)
            string[] newSourceRanges = new string[]
            {
                "Sheet1!A1:C10",   // existing range 1
                "Sheet2!A1:C10",   // existing range 2
                "Sheet3!A1:C10"    // newly added range
            };

            // Change the data source of the PivotTable to the new union range
            pivotTable.ChangeDataSource(newSourceRanges);

            // Refresh the PivotTable so it incorporates the new data
            pivotWorksheet.RefreshPivotTables();

            // Save the updated workbook
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
    }
}
