// Title: C# Example: Refresh an Aspose.Cells PivotTable after Changing Its ConsolidationFunction to Sum
// Common Searches: how to refresh a PivotTable after changing ConsolidationFunction using Aspose.Cells in C# | Aspose.Cells C# set PivotTable data field aggregation to Sum and recalculate | C# code to update PivotTable cache after modifying aggregation function with Aspose.Cells | refresh PivotTable data after changing ConsolidationFunction in .xlsx using Aspose.Cells | Aspose.Cells PivotTable RefreshData CalculateData example in C#
// Tags: Aspose.Cells PivotTable RefreshData method | Aspose.Cells PivotTable CalculateData method | C# set PivotTable ConsolidationFunction to Sum | update PivotTable aggregation function Aspose.Cells | recalculate PivotTable after data field change C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Provides PivotTable class

// The sample loads an existing workbook, accesses the first worksheet's PivotTable, sets each data field's ConsolidationFunction to Sum, calls RefreshData and CalculateData to update the cache, and saves the workbook.
class RefreshPivotTableExample
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the worksheet that contains the PivotTable (adjust index/name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (worksheet.PivotTables == null || worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTables found in the worksheet.");
                return;
            }

            // Get the first PivotTable
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Set the aggregation function of each data field to Sum
            foreach (PivotField dataField in pivotTable.DataFields)
            {
                dataField.Function = ConsolidationFunction.Sum;
            }

            // Refresh and recalculate the PivotTable
            pivotTable.RefreshData();   // Refreshes the source data cache
            pivotTable.CalculateData(); // Recalculates based on the new function

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
