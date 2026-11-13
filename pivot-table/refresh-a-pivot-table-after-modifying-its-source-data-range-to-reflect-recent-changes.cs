// Title: Refresh an Excel pivot table after modifying its source data using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to refresh pivot table after changing source cell values | programmatically recalculate Excel pivot table with Aspose.Cells after data update | Aspose.Cells RefreshData and CalculateData usage in .NET for pivot tables | update source range and refresh pivot table in Excel workbook using Aspose.Cells C# | example code to refresh pivot table after modifying data sheet with Aspose.Cells
// Tags: Aspose.Cells RefreshData method | Aspose.Cells CalculateData method | programmatic pivot table refresh C# | update source data range Excel pivot Aspose.Cells | refresh pivot after cell value change .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample loads a workbook, changes a cell value on the data sheet, accesses the first pivot table on another sheet, calls RefreshData() and CalculateData() to sync the pivot with the updated source, and saves the result to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "SourceData.xlsx";
            const string outputPath = "SourceData_Refreshed.xlsx";

            // Verify source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Retrieve required worksheets
            Worksheet dataSheet = workbook.Worksheets["Data"];
            Worksheet pivotSheet = workbook.Worksheets["Pivot"];

            if (dataSheet == null || pivotSheet == null)
            {
                Console.WriteLine("Required worksheets ('Data' or 'Pivot') are missing.");
                return;
            }

            // Example modification: change value in B5
            dataSheet.Cells["B5"].PutValue(2500);

            // Ensure a pivot table exists on the pivot sheet
            if (pivotSheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the 'Pivot' sheet.");
                return;
            }

            // Access the first pivot table
            PivotTable pivotTable = pivotSheet.PivotTables[0];

            // Refresh and recalculate the pivot table
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
