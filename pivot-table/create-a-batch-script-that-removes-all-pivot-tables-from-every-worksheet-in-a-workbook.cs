// Title: C# script to delete every pivot table from all worksheets in an Excel workbook using Aspose.Cells
// Common Searches: Aspose.Cells C# remove all pivot tables from workbook | how to programmatically delete pivot tables in each sheet with Aspose.Cells | C# code to clear pivot tables from an Excel file using Aspose.Cells | batch delete pivot tables across worksheets Aspose.Cells .NET | remove pivot tables from existing Excel workbook Aspose.Cells example
// Tags: Aspose.Cells pivot table deletion | C# iterate worksheets to clear pivots | Excel workbook pivot table cleanup using Aspose | programmatic removal of pivot tables .NET | batch pivot table removal Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel file (or creates a new one if missing), iterates through each worksheet, removes all pivot tables via the PivotTables collection, and saves the updated workbook while handling Aspose.Cells‑specific and general exceptions.
class RemovePivotTables
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            Workbook workbook;

            // Load existing workbook or create a new one if the file is missing
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one worksheet
                workbook.Worksheets[0].Name = "Sheet1";
            }

            // Remove all pivot tables from each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                for (int i = sheet.PivotTables.Count - 1; i >= 0; i--)
                {
                    sheet.PivotTables.RemoveAt(i);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (CellsException ex)
        {
            // Handles Aspose.Cells specific errors (e.g., corrupted file)
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handles all other runtime errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
