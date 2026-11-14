// Title: Refresh a PivotTable after adding items to a slicer cache using Aspose.Cells for .NET (C#)
// Common Searches: how to refresh a pivot table after updating slicer cache with Aspose.Cells C# | Aspose.Cells C# refresh pivot table linked to slicer cache | verify that slicer cache changes are reflected in pivot table using Aspose.Cells | C# code to refresh pivot table after modifying slicer in workbook with Aspose.Cells | Aspose.Cells refresh pivot data after slicer cache modification
// Tags: Aspose.Cells pivot table refresh after slicer cache update | C# refresh pivot table linked to slicer using Aspose.Cells | validate slicer cache changes reflected in pivot table Aspose.Cells | save workbook after pivot refresh Aspose.Cells C# | Aspose.Cells slicer cache manipulation example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;          // Required for PivotTable class
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;

// The example loads a workbook, checks for a slicer and a pivot table on the first worksheet, refreshes the pivot table with RefreshData and CalculateData to reflect new slicer cache items, optionally confirms the slicer is linked to a cache, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Please ensure the file exists.");
                return;
            }

            // Load the workbook that contains the slicer and pivot table
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one slicer
            if (sheet.Slicers.Count == 0)
            {
                Console.WriteLine("No slicers found on the first worksheet.");
                return;
            }

            // Get the first slicer on the worksheet
            Slicer slicer = sheet.Slicers[0];
            Console.WriteLine($"Slicer \"{slicer.Name}\" found.");

            // Attempt to locate a pivot table on the same worksheet
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the first worksheet.");
                return;
            }

            // Retrieve the first pivot table
            PivotTable pivotTable = sheet.PivotTables[0];
            Console.WriteLine($"Pivot table \"{pivotTable.Name}\" found.");

            // Refresh the pivot table to ensure it reflects the current data
            try
            {
                pivotTable.RefreshData();
                pivotTable.CalculateData();
                Console.WriteLine("Pivot table refreshed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to refresh pivot table: {ex.Message}");
            }

            // Optional: Verify that the slicer is linked to a slicer cache
            if (slicer.SlicerCache != null)
            {
                Console.WriteLine("Slicer is linked to a slicer cache.");
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved as \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
