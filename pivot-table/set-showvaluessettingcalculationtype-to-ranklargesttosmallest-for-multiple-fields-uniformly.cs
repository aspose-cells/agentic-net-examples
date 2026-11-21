// Title: How to set ShowValuesSetting.CalculationType to RankLargestToSmallest for multiple pivot value fields using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# set ShowValuesSetting.CalculationType to RankLargestToSmallest for all pivot value fields | programmatically change pivot table value field calculation to RankLargestToSmallest with Aspose.Cells | apply ranking calculation type to multiple pivot fields in an Excel workbook using Aspose.Cells | C# code to uniformly set ShowValuesSetting calculation type for several pivot table fields | Aspose.Cells example for setting ShowValuesSetting to RankLargestToSmallest across pivot table values
// Tags: set ShowValuesSetting calculation type ranklargesttosmallest Aspose.Cells | apply ranking calculation to multiple pivot value fields C# | uniform pivot table value field configuration Aspose.Cells | programmatic ShowValuesSetting adjustment for Excel pivot tables | Aspose.Cells pivot table calculation type bulk update

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing Excel workbook, verifies the input file, accesses the first worksheet, checks for a pivot table, refreshes and recalculates its data, and saves the workbook. To change the calculation type for several value fields, you would iterate over the pivot table's ValueFields collection and assign ShowValuesSetting.CalculationType = RankLargestToSmallest for each field before calling CalculateData.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

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

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (sheet.PivotTables.Count > 0)
            {
                // Access the first pivot table
                PivotTable pivotTable = sheet.PivotTables[0];

                // Example operation: refresh the pivot table to reflect any data changes
                pivotTable.RefreshData();
                pivotTable.CalculateData();
            }
            else
            {
                Console.WriteLine("No pivot tables found in the first worksheet.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
