// Title: Remove empty rows from the first worksheet of an Excel workbook using Aspose.Cells for .NET with DeleteOptions (UpdateReference = false)
// Common Searches: C# Aspose.Cells delete blank rows without updating formulas | How to use DeleteOptions.UpdateReference false to remove empty rows in .NET | Aspose.Cells DeleteBlankRows on first worksheet example | Remove empty rows from Excel file using Aspose.Cells API | Delete blank rows in a workbook while preserving cell references in C#
// Tags: Aspose.Cells DeleteBlankRows API | DeleteOptions UpdateReference false .NET | remove empty rows first worksheet Excel | C# Excel blank row deletion Aspose | Excel workbook row cleanup Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel file (or creates a new workbook if the file is missing), accesses the first worksheet, and deletes all blank rows using Cells.DeleteBlankRows with DeleteOptions that have UpdateReference set to false, then saves the workbook, handling both Aspose.Cells‑specific and general exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        Workbook workbook = null;

        try
        {
            // Verify that the input file exists before attempting to load it
            if (File.Exists(inputPath))
            {
                // Load the workbook; if it is password‑protected, an exception will be thrown
                workbook = new Workbook(inputPath);
            }
            else
            {
                // Create a new workbook with a default worksheet
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Delete blank rows without updating references
            DeleteOptions options = new DeleteOptions
            {
                UpdateReference = false
            };
            sheet.Cells.DeleteBlankRows(options);

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (CellsException ex)
        {
            // Handle Aspose.Cells specific errors (e.g., password‑protected files)
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
