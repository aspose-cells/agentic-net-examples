// Title: Refresh an Aspose.Cells pivot table after updating source worksheet cells in C#
// Common Searches: Aspose.Cells C# refresh pivot table after modifying source range | programmatically update pivot cache in Aspose.Cells after cell change | C# code to recalculate pivot table when source data changes using Aspose.Cells | how to call RefreshData and CalculateData on Aspose.Cells pivot table | update Excel pivot table with Aspose.Cells after editing worksheet values
// Tags: Aspose.Cells RefreshData API | Aspose.Cells CalculateData method | C# pivot table source range update | Aspose.Cells modify worksheet cell then refresh pivot | Excel pivot cache refresh using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Ensure PivotTable class is available

// The example loads a workbook, changes a cell in the source worksheet, retrieves the first pivot table on another sheet, invokes RefreshData() and CalculateData() to synchronize the pivot with the new data, and saves the updated file, demonstrating error handling throughout.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "SourceData.xlsx";
            const string outputPath = "Updated.xlsx";

            // Verify source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source file not found: {sourcePath}");

            // Load the workbook containing source data and the pivot table
            var workbook = new Workbook(sourcePath);

            // -------------------------------------------------
            // Modify the underlying source data range
            // -------------------------------------------------
            // Example: change the value in cell A2 of the source worksheet named "Data"
            Worksheet sourceSheet = workbook.Worksheets["Data"];
            if (sourceSheet == null)
                throw new InvalidOperationException("Worksheet 'Data' not found.");

            sourceSheet.Cells["A2"].PutValue(12345); // put your new value here

            // -------------------------------------------------
            // Refresh the pivot table to reflect the changed data
            // -------------------------------------------------
            // Assume the pivot table resides on a worksheet named "Pivot"
            Worksheet pivotSheet = workbook.Worksheets["Pivot"];
            if (pivotSheet == null)
                throw new InvalidOperationException("Worksheet 'Pivot' not found.");

            // Get the first pivot table on that sheet
            if (pivotSheet.PivotTables.Count == 0)
                throw new InvalidOperationException("No pivot tables found on the 'Pivot' sheet.");

            PivotTable pivotTable = pivotSheet.PivotTables[0];

            // Refresh the data source of the pivot table
            pivotTable.RefreshData();

            // Recalculate the pivot table (optional but ensures values are updated)
            pivotTable.CalculateData();

            // -------------------------------------------------
            // Save the workbook with the refreshed pivot table
            // -------------------------------------------------
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log or display the error details
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
