// Title: Refresh an Excel pivot table after modifying its source data using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# refresh pivot table after updating source cells | programmatically recalculate Excel pivot table with Aspose.Cells .NET | how to call RefreshData and CalculateData on a pivot table in C# | update Excel pivot table data range using Aspose.Cells workbook API | C# example to modify source data and refresh pivot table in Aspose.Cells
// Tags: Aspose.Cells RefreshData method | Aspose.Cells CalculateData function | C# programmatic Excel pivot refresh | modify Excel source data with Aspose.Cells | update pivot table after data change .NET

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// Loads a workbook, changes a cell in the source worksheet, accesses the first pivot table on another sheet, calls RefreshData() and CalculateData() to sync the pivot with the updated data, and saves the result to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook containing the source data and the pivot table
            Workbook workbook = new Workbook(inputPath);

            // Access the worksheet that holds the source data (adjust the name as needed)
            Worksheet dataSheet = workbook.Worksheets["Data"];
            if (dataSheet == null)
            {
                Console.WriteLine("Error: Worksheet \"Data\" not found.");
                return;
            }

            // Example modification: change the value of a cell within the source range
            dataSheet.Cells["B2"].PutValue(12345);

            // Access the worksheet that contains the pivot table (adjust the name as needed)
            Worksheet pivotSheet = workbook.Worksheets["Pivot"];
            if (pivotSheet == null)
            {
                Console.WriteLine("Error: Worksheet \"Pivot\" not found.");
                return;
            }

            // Ensure there is at least one pivot table on the worksheet
            if (pivotSheet.PivotTables.Count == 0)
            {
                Console.WriteLine("Error: No pivot tables found on the \"Pivot\" worksheet.");
                return;
            }

            // Get the first pivot table on the worksheet
            PivotTable pivotTable = pivotSheet.PivotTables[0];

            // Refresh the pivot table to pick up the updated source data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook with the refreshed pivot table
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
