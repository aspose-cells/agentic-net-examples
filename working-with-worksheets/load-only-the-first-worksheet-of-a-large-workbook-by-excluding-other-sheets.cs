// Title: Load a large Excel workbook and retain only the first worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, deletes every worksheet except the first, and saves the trimmed workbook to a new file. | Show how to efficiently trim a loaded workbook to a single sheet in .NET by removing extra worksheets with Aspose.Cells. | Provide a C# snippet that verifies the source file, loads the workbook, removes all but the first worksheet, and writes the result to a separate file.
// Common Searches: Aspose.Cells .NET load workbook and delete all sheets except the first | C# remove extra worksheets from large Excel file using Aspose.Cells | how to save only the first worksheet of an Excel workbook with Aspose.Cells for .NET | optimize memory when loading large Excel workbook by discarding unnecessary sheets in C#
// Tags: Aspose.Cells workbook trim to first sheet | C# delete trailing worksheets Aspose.Cells | load large Excel file with single worksheet .NET | remove worksheets by index Aspose.Cells API

using System;
using System.IO;
using Aspose.Cells;

// Loads 'largeWorkbook.xlsx', removes every worksheet except the first, and saves the resulting single‑sheet workbook as 'firstSheetOnly.xlsx' using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputPath = "largeWorkbook.xlsx";
        const string outputPath = "firstSheetOnly.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the entire workbook
            Workbook workbook = new Workbook(inputPath);

            // Keep only the first worksheet and remove the rest
            while (workbook.Worksheets.Count > 1)
            {
                // Remove the last worksheet (index is zero‑based)
                workbook.Worksheets.RemoveAt(workbook.Worksheets.Count - 1);
            }

            // Save the workbook (contains only the first worksheet)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
