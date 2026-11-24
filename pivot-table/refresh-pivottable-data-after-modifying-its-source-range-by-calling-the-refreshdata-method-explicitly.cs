// Title: C# example: Refresh an Aspose.Cells PivotTable after changing its source data range
// Common Searches: c# aspnet refresh pivot table after editing source data using Aspose.Cells | how to use RefreshData on an Aspose.Cells PivotTable in .NET | example of updating a cell value and refreshing the pivot cache with Aspose.Cells C# | programmatically refresh pivot table cache after changing source range in a workbook | Aspose.Cells C# code to refresh pivot table after modifying source worksheet
// Tags: Aspose.Cells RefreshData pivot table | C# refresh pivot cache Aspose.Cells | update source range Aspose.Cells PivotTable | programmatic pivot table refresh .NET | modify cell and refresh pivot Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable class

// The sample loads a workbook, changes a cell in the data sheet, accesses the first PivotTable on the pivot sheet, calls RefreshData to update the pivot cache, and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        // Define file paths (replace with actual paths as needed)
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Update a cell in the data sheet (example: B2)
            Worksheet dataSheet = workbook.Worksheets["Data"];
            if (dataSheet == null)
                throw new InvalidOperationException("Worksheet named 'Data' was not found.");

            dataSheet.Cells["B2"].PutValue(12345);

            // Access the pivot sheet and its first pivot table
            Worksheet pivotSheet = workbook.Worksheets["Pivot"];
            if (pivotSheet == null)
                throw new InvalidOperationException("Worksheet named 'Pivot' was not found.");

            if (pivotSheet.PivotTables == null || pivotSheet.PivotTables.Count == 0)
                throw new InvalidOperationException("No pivot tables found on the 'Pivot' sheet.");

            PivotTable pivotTable = pivotSheet.PivotTables[0];

            // Refresh the pivot table data
            pivotTable.RefreshData();

            // Save the modified workbook
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
