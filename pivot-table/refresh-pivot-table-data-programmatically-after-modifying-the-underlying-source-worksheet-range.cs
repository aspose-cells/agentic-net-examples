// Title: Refresh all pivot tables after updating source worksheet data with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that changes a cell value in a source worksheet, then iterates through every PivotTable in the workbook to call RefreshData and CalculateData using Aspose.Cells. | Generate a robust Aspose.Cells example that loads an Excel file, validates the presence of specific worksheets, updates source data, refreshes all pivot tables, and saves the result with error handling. | Create a C# snippet that programmatically refreshes pivot table caches after modifying the underlying range, including handling missing files and worksheets.
// Common Searches: Aspose.Cells C# refresh pivot table cache after editing source cells | How to programmatically update pivot tables in an Excel workbook using Aspose.Cells for .NET | C# example to call RefreshData and CalculateData on all pivot tables in a workbook | Error handling for missing worksheets when refreshing pivot tables with Aspose.Cells
// Tags: Aspose.Cells RefreshData CalculateData pivot tables | C# update source worksheet cell Aspose.Cells | programmatic pivot cache refresh Excel .NET | validate worksheet existence Aspose.Cells | save workbook after pivot refresh Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable classes

// The example loads an existing workbook, updates a cell in the 'SourceData' worksheet, verifies required worksheets, iterates through each PivotTable on the 'PivotSheet' worksheet calling RefreshData and CalculateData, and then saves the modified workbook. It includes checks for file existence and missing worksheets to ensure reliable pivot cache refresh.
class RefreshPivotTableExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "OutputWorkbook.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the workbook that contains the source data and the pivot table
            Workbook workbook = new Workbook(inputPath);

            // -------------------------------------------------
            // Modify the underlying source data range
            // -------------------------------------------------
            // Assume the source data is on a worksheet named "SourceData"
            Worksheet sourceSheet = workbook.Worksheets["SourceData"];
            if (sourceSheet == null)
                throw new InvalidOperationException("Worksheet 'SourceData' does not exist.");

            // Example: change the value in cell B2
            sourceSheet.Cells["B2"].PutValue(12345);
            // Additional modifications can be performed here

            // -------------------------------------------------
            // Refresh all pivot tables that depend on the source data
            // -------------------------------------------------
            // Assume the pivot table is on a worksheet named "PivotSheet"
            Worksheet pivotSheet = workbook.Worksheets["PivotSheet"];
            if (pivotSheet == null)
                throw new InvalidOperationException("Worksheet 'PivotSheet' does not exist.");

            PivotTableCollection pivotTables = pivotSheet.PivotTables;

            foreach (PivotTable pivotTable in pivotTables)
            {
                // Refresh the pivot table's data cache from the modified source range
                pivotTable.RefreshData();

                // Recalculate the pivot table (e.g., totals, subtotals, calculated fields)
                pivotTable.CalculateData();

                // Note: RefreshDataOnLoad property is not available in this version of Aspose.Cells.
                // The pivot table will reflect the refreshed data after saving.
            }

            // -------------------------------------------------
            // Save the workbook with the refreshed pivot table
            // -------------------------------------------------
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log or display the error details for troubleshooting
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
