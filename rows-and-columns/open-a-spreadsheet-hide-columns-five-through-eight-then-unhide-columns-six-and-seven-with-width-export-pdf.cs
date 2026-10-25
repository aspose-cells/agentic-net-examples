// Title: Hide columns 5‑8, unhide columns 6‑7 with custom width, and export the worksheet to PDF using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells hide a range of columns then set width for specific columns before PDF export | How to unhide columns with custom width in Aspose.Cells and save workbook as PDF | Aspose.Cells hide columns 5 to 8 and export the sheet to PDF in .NET
// Tags: hide column range Aspose.Cells C# | unhide column with width Aspose.Cells | set column width Aspose.Cells PDF export | export worksheet to PDF Aspose.Cells | column visibility manipulation Aspose.Cells .NET

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing XLSX file, hides columns 5 through 8, then unhides columns 6 and 7 while applying a custom width, and finally saves the worksheet as a PDF, including file‑existence checks and exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing spreadsheet
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Hide columns 5 through 8 (1‑based column numbers correspond to indexes 4‑7)
            for (int col = 4; col <= 7; col++)
            {
                sheet.Cells.HideColumn(col);
            }

            // Unhide columns 6 and 7 (indexes 5 and 6) and set a specific width
            int[] colsToUnhide = { 5, 6 };
            foreach (int col in colsToUnhide)
            {
                // Provide a width when unhiding; 0 uses the default column width
                sheet.Cells.UnhideColumn(col, 0);
                // Set column width (in characters). Adjust the value as required.
                sheet.Cells.SetColumnWidth(col, 15);
            }

            // Export the workbook (or the active sheet) to PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
