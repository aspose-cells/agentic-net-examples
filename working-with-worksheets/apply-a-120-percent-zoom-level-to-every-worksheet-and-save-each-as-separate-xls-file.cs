// Title: How to set a 120% zoom on every worksheet and save each sheet as an individual XLS file with Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells in C# to iterate through all worksheets, set each sheet’s Zoom property to 120, and write the sheet to its own Excel 97‑2003 (.xls) file. | Create a separate workbook for each worksheet, copy the original sheet, apply a 120 % zoom level, and save the workbook as an .xls file while sanitizing the sheet name for a valid filename.
// Common Searches: Aspose.Cells set worksheet zoom to 120 percent in C# | Save each worksheet as a separate .xls file using Aspose.Cells .NET | Copy a sheet to a new workbook and apply zoom with Aspose.Cells | How to generate individual Excel 97‑2003 files from a multi‑sheet workbook in C# | Sanitize worksheet names for file output Aspose.Cells
// Tags: worksheet zoom 120% Aspose.Cells | export single sheet to XLS Aspose.Cells | copy worksheet to new workbook C# | sanitize filename from sheet name C# | Aspose.Cells SaveFormat.Excel97To2003 | iterate worksheets Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

// The program loads a multi‑sheet workbook, sets each worksheet’s Zoom property to 120 %, creates a new workbook containing only that sheet, sanitizes the sheet name to form a valid filename, and saves each sheet as an individual Excel 97‑2003 (.xls) file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(inputPath);

            // Process each worksheet
            foreach (Worksheet sheet in sourceWorkbook.Worksheets)
            {
                // Apply 120% zoom
                sheet.Zoom = 120;

                // Create a new workbook for the single sheet
                Workbook singleSheetWorkbook = new Workbook();

                // Remove the default empty sheet
                singleSheetWorkbook.Worksheets.Clear();

                // Copy the current worksheet by name
                singleSheetWorkbook.Worksheets.AddCopy(sheet.Name);

                // Rename the copied sheet to match the original name
                singleSheetWorkbook.Worksheets[0].Name = sheet.Name;

                // Ensure a valid file name for the output
                string safeName = MakeValidFileName(sheet.Name);
                string outputFileName = $"{safeName}.xls";

                // Save as Excel 97‑2003 format
                singleSheetWorkbook.Save(outputFileName, SaveFormat.Excel97To2003);
                Console.WriteLine($"Saved sheet '{sheet.Name}' to {outputFileName}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper method to replace invalid filename characters
    private static string MakeValidFileName(string name)
    {
        foreach (char c in Path.GetInvalidFileNameChars())
        {
            name = name.Replace(c, '_');
        }
        return name;
    }
}
