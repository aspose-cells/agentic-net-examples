// Title: Convert an Excel workbook to a static PDF and strip all hyperlinks using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells remove all hyperlinks from Excel before saving as PDF | How to export an .xlsx to a printable PDF without active links using Aspose.Cells | Aspose.Cells C# clear worksheet hyperlinks then convert workbook to PDF | Generate a static PDF from Excel file with hyperlinks disabled in .NET | Aspose.Cells PDF export without clickable hyperlinks in C#
// Tags: Aspose.Cells remove hyperlinks workbook | Excel to PDF conversion static Aspose.Cells | C# clear worksheet hyperlinks Aspose.Cells | SaveFormat.Pdf export without links | printable PDF generation from .xlsx Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the input XLSX file, loads it into an Aspose.Cells Workbook, clears all hyperlinks from each worksheet, and then saves the workbook as a static PDF using SaveFormat.Pdf, with error handling for Aspose.Cells-specific and general exceptions.
class WorkbookToPdfStatic
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Remove all hyperlinks from each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.Hyperlinks.Clear();
            }

            // Save the modified workbook as a PDF (static, without active hyperlinks)
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (CellsException ex)
        {
            // Handles Aspose.Cells specific errors (e.g., password-protected files)
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handles any other unexpected errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
