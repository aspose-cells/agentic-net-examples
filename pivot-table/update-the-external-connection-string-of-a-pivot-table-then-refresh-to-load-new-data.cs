// Title: Refresh a pivot table’s external data source and recalculate its values with Aspose.Cells for .NET (C#)
// Common Searches: how to refresh pivot table cache after changing data source using Aspose.Cells in C# | C# Aspose.Cells recalculate pivot table after updating external connection | load Excel workbook and refresh first pivot table with Aspose.Cells .NET | Aspose.Cells refresh pivot data and save workbook programmatically | C# example for handling missing pivot tables when refreshing data with Aspose.Cells
// Tags: Aspose.Cells refresh pivot cache C# | recalculate pivot table Aspose.Cells .NET | load workbook update pivot data source Aspose | handle missing pivot tables Aspose.Cells | save workbook after pivot refresh Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample loads an existing Excel file, accesses the first worksheet’s first pivot table, attempts to refresh its external data cache, recalculates the pivot values, and saves the workbook to a new file, while providing console messages for missing files or absent pivot tables.
class UpdatePivotExternalConnection
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
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Assume the pivot table is on the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("Error: No pivot tables found on the first worksheet.");
                return;
            }

            // Get the first pivot table in the worksheet
            PivotTable pivot = worksheet.PivotTables[0];

            // Refresh the pivot cache (loads data from the existing connection)
            try
            {
                pivot.RefreshData();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Unable to refresh pivot data. {ex.Message}");
            }

            // Recalculate the pivot table after refresh
            try
            {
                pivot.CalculateData();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Unable to calculate pivot data. {ex.Message}");
            }

            // Save the workbook with the updated pivot table
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
