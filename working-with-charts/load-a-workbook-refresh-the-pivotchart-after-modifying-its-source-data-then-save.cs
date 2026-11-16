// Title: Refresh a PivotChart after modifying source data and save the workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# code to refresh a PivotChart after updating source cells with Aspose.Cells | How to programmatically refresh a PivotTable and its chart in an Excel file using Aspose.Cells .NET | Aspose.Cells example for modifying worksheet data then refreshing the associated pivot chart before saving | Refresh all pivot tables in a workbook with Aspose.Cells C# | Update a cell value and refresh the linked pivot chart in Excel using the Aspose.Cells library
// Tags: aspnet pivot chart refresh using Aspose.Cells | update worksheet cell value with Aspose.Cells C# | programmatic pivot table data refresh .NET | save Excel workbook after pivot refresh Aspose.Cells | load workbook modify source data Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// This C# example loads an existing Excel workbook, changes a cell in the source data sheet, accesses the first PivotTable, calls RefreshData to update the linked PivotChart (or all PivotTables), and then saves the workbook to a new file, with error handling for missing files or worksheets.
class PivotChartRefreshExample
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

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the worksheet that contains the source data (replace with actual sheet name if needed)
            Worksheet dataSheet = workbook.Worksheets["Data"];
            if (dataSheet == null)
                throw new InvalidOperationException("Worksheet named 'Data' does not exist in the workbook.");

            // Modify source data (example: update cell A2)
            dataSheet.Cells["A2"].PutValue(12345);

            // Access the worksheet that contains the PivotTable (assumed to be the first worksheet)
            Worksheet pivotSheet = workbook.Worksheets[0];
            if (pivotSheet.PivotTables.Count == 0)
                throw new InvalidOperationException("No PivotTables found on the first worksheet.");

            // Get the first PivotTable
            PivotTable pivotTable = pivotSheet.PivotTables[0];

            // Refresh the PivotTable to reflect the changed source data
            pivotTable.RefreshData();

            // Optionally, refresh all PivotTables in the workbook
            //foreach (Worksheet ws in workbook.Worksheets)
            //{
            //    foreach (PivotTable pt in ws.PivotTables)
            //        pt.RefreshData();
            //}

            // Save the workbook with the refreshed PivotChart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log or display the error details
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
