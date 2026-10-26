// Title: Delete a pivot table at a specific zero‑based index using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# remove pivot table by its zero based index | Worksheet.PivotTables.RemoveAt example code for .NET | How to delete the first pivot table in an Excel worksheet with Aspose.Cells | Check pivot table count before calling RemoveAt in C# | Programmatically delete a pivot table and save the workbook using Aspose.Cells
// Tags: Aspose.Cells pivot table removal by index | Worksheet.PivotTables.RemoveAt usage | Excel workbook pivot table deletion .NET | validate pivot table index Aspose.Cells | remove pivot table and save workbook C#

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel file (or creates a new workbook if missing), accesses the first worksheet, verifies that a zero‑based pivot table index is within the collection bounds, removes the pivot table at that index with Worksheet.PivotTables.RemoveAt, and saves the modified workbook, handling Aspose.Cells‑specific and generic exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        Workbook workbook = null;

        try
        {
            if (File.Exists(inputPath))
            {
                // Load the workbook. If the file is password‑protected, set the password in LoadOptions.
                var loadOptions = new LoadOptions();
                // loadOptions.Password = "yourPassword"; // uncomment and set if needed
                workbook = new Workbook(inputPath, loadOptions);
            }
            else
            {
                // Create a new workbook when the input file is missing.
                workbook = new Workbook();
                workbook.Worksheets.Add("Sheet1");
            }

            // Get the first worksheet (adjust index or name as required).
            Worksheet worksheet = workbook.Worksheets[0];

            // Zero‑based index of the pivot table to delete.
            int pivotIndex = 0;

            // Delete the pivot table if the index is valid.
            if (pivotIndex >= 0 && pivotIndex < worksheet.PivotTables.Count)
            {
                worksheet.PivotTables.RemoveAt(pivotIndex);
            }

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (CellsException ex)
        {
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
