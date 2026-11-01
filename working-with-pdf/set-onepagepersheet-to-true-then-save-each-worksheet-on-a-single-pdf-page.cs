// Title: How to export every Excel worksheet to its own single‑page PDF with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx workbook, loops through all worksheets, and saves each one as an individual PDF file using PdfSaveOptions.OnePagePerSheet = true. | Demonstrate how to copy a single worksheet into a new temporary Workbook, clean the worksheet name for a safe file name, and export it to PDF with Aspose.Cells. | Create a robust C# routine that verifies the input Excel file exists, handles exceptions for each sheet during PDF conversion, and names the output PDFs after their corresponding worksheets.
// Common Searches: Aspose.Cells C# export each sheet to separate PDF one page per sheet | PdfSaveOptions OnePagePerSheet usage example .NET | How to save a single Excel worksheet as a PDF with Aspose.Cells | C# replace invalid filename characters when generating PDF from Excel sheets
// Tags: export worksheet to PDF Aspose.Cells | PdfSaveOptions OnePagePerSheet .NET | save each Excel sheet as separate PDF | copy worksheet to new workbook Aspose.Cells | sanitize filename characters C#

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an input.xlsx workbook, iterates through its worksheets, creates a temporary workbook containing only the current sheet, sets PdfSaveOptions.OnePagePerSheet to true, sanitizes the sheet name for a safe file name, and saves each sheet as a separate single‑page PDF while handling missing file and per‑sheet export errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file not found – {inputPath}");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and save each one as a separate PDF page
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                try
                {
                    // Create a temporary workbook containing only the current worksheet
                    Workbook sheetWb = new Workbook();
                    sheetWb.Worksheets.Clear();

                    // Add a copy of the current worksheet by name
                    sheetWb.Worksheets.AddCopy(workbook.Worksheets[i].Name);

                    // Configure PDF save options for the current sheet
                    PdfSaveOptions pdfOptions = new PdfSaveOptions
                    {
                        OnePagePerSheet = true // Force one page per sheet
                    };

                    // Build a safe file name for the PDF (e.g., Sheet1_SheetName.pdf)
                    string safeSheetName = workbook.Worksheets[i].Name.Replace(Path.GetInvalidFileNameChars(), '_');
                    string pdfFileName = $"Sheet{i + 1}_{safeSheetName}.pdf";

                    // Save the selected worksheet to PDF using the configured options
                    sheetWb.Save(pdfFileName, pdfOptions);
                    Console.WriteLine($"Saved: {pdfFileName}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save sheet {i + 1}: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Extension method to replace invalid filename characters
static class StringExtensions
{
    public static string Replace(this string str, char[] chars, char replacement)
    {
        foreach (char c in chars)
        {
            str = str.Replace(c, replacement);
        }
        return str;
    }
}
