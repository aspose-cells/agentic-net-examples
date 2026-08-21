// Title: Aspose.Cells for .NET – Disable Blank Page Generation When Saving an Empty Worksheet to PDF
// Description: Demonstrates how to prevent Aspose.Cells from adding a blank page when a workbook has no printable content. The example hides the default sheet, adds an empty visible worksheet, sets a non‑printable print area, and saves the workbook as PDF with PdfSaveOptions.OutputBlankPageWhenNothingToPrint set to false.
// Keywords: Aspose.Cells PdfSaveOptions | OutputBlankPageWhenNothingToPrint | disable blank page PDF | save empty worksheet to PDF | C# Aspose.Cells PDF conversion | hide worksheet Aspose.Cells | print area no printable cells
// Common Searches: Aspose.Cells prevent blank page in PDF | PdfSaveOptions disable blank page when nothing to print | C# save empty Excel sheet as PDF without extra page | How to hide default worksheet in Aspose.Cells | Set print area to A1:A0 Aspose.Cells
// Developer Intent: Avoid generating a blank PDF page when the workbook contains no printable cells.
// Use Cases: Creating clean PDF reports where some worksheets are intentionally empty. | Automating Excel‑to‑PDF pipelines that must not include placeholder pages. | Generating printable PDFs from dynamic workbooks with optional data sections.
// AI Prompts: Show C# code that configures Aspose.Cells PdfSaveOptions to skip blank pages when the print area is empty. | Explain how OutputBlankPageWhenNothingToPrint works and when to use it in PDF conversion. | Provide a step‑by‑step guide to hide the default sheet, add an empty visible sheet, set a non‑printable area, and save to PDF without a blank page using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to prevent Aspose.Cells from adding a blank page when a workbook has no printable content. The example hides the default sheet, adds an empty visible worksheet, sets a non‑printable print area, and saves the workbook as PDF with PdfSaveOptions.OutputBlankPageWhenNothingToPrint set to false.
class DisableBlankPagePdf
{
    static void Main()
    {
        try
        {
            // Create a new (empty) workbook
            Workbook workbook = new Workbook();

            // Ensure there is at least one visible worksheet.
            // Hide the default worksheet and add a new empty visible worksheet.
            Worksheet defaultSheet = workbook.Worksheets[0];
            defaultSheet.IsVisible = false;

            // Add a new worksheet and obtain its reference
            int newSheetIndex = workbook.Worksheets.Add();
            Worksheet emptySheet = workbook.Worksheets[newSheetIndex];

            // Set a print area that results in no printable cells.
            emptySheet.PageSetup.PrintArea = "A1:A0";

            // Configure PDF save options: do NOT generate a blank page when nothing is printable
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OutputBlankPageWhenNothingToPrint = false
            };

            // Define the output file path (e.g., on the desktop)
            string outputPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "NoBlankPage.pdf");

            // Ensure the directory exists
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
