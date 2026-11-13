// Title: C# Example: Refresh an Aspose.Cells PivotTable after Changing Its ConsolidationFunction to Average
// Common Searches: Aspose.Cells C# refresh pivot table after setting ConsolidationFunction to Average | How to update aggregated results in an Excel PivotTable using Aspose.Cells RefreshData | C# code to change pivot data field function and recalculate workbook with Aspose.Cells | RefreshData method usage for Aspose.Cells PivotTable in .NET | Update pivot cache after modifying ConsolidationFunction in Aspose.Cells C#
// Tags: Aspose.Cells RefreshData pivot table | C# change PivotTable ConsolidationFunction | update Excel pivot aggregation Aspose.Cells | recalculate workbook formulas after pivot refresh | load and save workbook with modified pivot C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing Excel workbook, accesses its first PivotTable, sets the first data field's ConsolidationFunction to Average, calls the RefreshData method to refresh the pivot cache, recalculates all formulas, and saves the updated workbook to a new file.
class PivotTableRefreshExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (worksheet.PivotTables == null || worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTable found in the worksheet.");
                return;
            }

            // Get the first PivotTable (or locate by name)
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Change the aggregation function of the first data field to Average
            if (pivotTable.DataFields != null && pivotTable.DataFields.Count > 0)
            {
                pivotTable.DataFields[0].Function = ConsolidationFunction.Average;
            }
            else
            {
                Console.WriteLine("PivotTable has no data fields to modify.");
                return;
            }

            // Refresh the PivotTable data (using the available RefreshData method)
            try
            {
#pragma warning disable CS0618 // Suppress obsolete warning for RefreshData
                pivotTable.RefreshData();
#pragma warning restore CS0618
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error refreshing pivot table: {ex.Message}");
                return;
            }

            // Recalculate workbook formulas after refresh
            workbook.CalculateFormula();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
