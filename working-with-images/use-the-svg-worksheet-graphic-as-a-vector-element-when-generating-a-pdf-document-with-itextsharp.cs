// Title: Convert only the first worksheet of an Excel file to a single‑page PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx workbook, hides every sheet except the first, sets the first sheet to fit one page width, and saves it as a PDF with Aspose.Cells. | Show how to configure PdfSaveOptions in Aspose.Cells to generate one PDF page per worksheet while preserving the original page layout. | Create a reusable C# method that accepts input and output paths, applies the necessary page‑setup settings to the first worksheet, and exports it to PDF with proper error handling.
// Common Searches: Aspose.Cells C# export first worksheet to PDF single page | fit worksheet to one page width when saving as PDF using Aspose.Cells | hide all sheets except first before PDF conversion Aspose.Cells .NET | PdfSaveOptions OnePagePerSheet example C# Aspose.Cells | set page setup fit to pages wide Aspose.Cells PDF export
// Tags: Aspose.Cells first worksheet PDF export | fit worksheet to single page width Aspose.Cells | hide Excel sheets before PDF conversion Aspose.Cells | PdfSaveOptions OnePagePerSheet configuration | C# Excel to PDF page setup Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, hides all worksheets except the first, configures the first sheet to fit one page width, applies PdfSaveOptions with OnePagePerSheet, and saves the result as a PDF, including basic file‑existence checks and exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input Excel file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Hide all worksheets except the first one to export only the first sheet
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                if (i != 0)
                {
                    workbook.Worksheets[i].IsVisible = false; // hide worksheet
                }
            }

            // Configure page setup for the first worksheet (fit to one page width)
            Worksheet firstSheet = workbook.Worksheets[0];
            firstSheet.PageSetup.FitToPagesWide = 1;
            firstSheet.PageSetup.FitToPagesTall = 0;

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
            };

            // Save the first worksheet as PDF
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF successfully created at: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
