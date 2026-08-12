// Title: C# – Validate that all columns fit on a single PDF page per worksheet with Aspose.Cells
// Description: Loads an Excel workbook, sets each worksheet’s FitToPagesWide to 1, enables AllColumnsInOnePagePerSheet and OnePagePerSheet, renders each sheet to verify the page count, logs a warning for sheets that exceed one page, and saves the workbook as a PDF with the same one‑page‑wide configuration.
// Keywords: Aspose.Cells | C# | .NET | FitToPagesWide | AllColumnsInOnePagePerSheet | PDF conversion | column width validation | SheetRender page count | PdfSaveOptions | Excel to PDF
// Common Searches: Aspose.Cells fit all columns on one PDF page | C# validate worksheet page count before PDF export | How to use AllColumnsInOnePagePerSheet with Aspose.Cells | Set FitToPagesWide = 1 for each sheet in Aspose.Cells | Check if Excel columns fit on a single PDF page .NET
// Developer Intent: Confirm that every worksheet renders to exactly one PDF page and enforce the setting before saving.
// Use Cases: Validate column fitting prior to generating PDF reports to avoid unintended multi‑page spreads. | Automatically adjust PageSetup for each worksheet so the PDF output is one page wide. | Detect and log worksheets that still exceed one page after applying FitToPagesWide = 1.
// AI Prompts: Show C# code that uses Aspose.Cells to verify each worksheet renders to a single page when AllColumnsInOnePagePerSheet is enabled. | Explain how to programmatically set FitToPagesWide = 1 for all worksheets and save the workbook as a one‑page‑wide PDF. | Provide a method to capture and handle cases where a worksheet exceeds one page after applying the one‑page‑wide setting.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an Excel workbook, sets each worksheet’s FitToPagesWide to 1, enables AllColumnsInOnePagePerSheet and OnePagePerSheet, renders each sheet to verify the page count, logs a warning for sheets that exceed one page, and saves the workbook as a PDF with the same one‑page‑wide configuration.
class FitAllColumnsValidator
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists; create a new one if it does not.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                workbook = new Workbook();
                // Optionally add a sample worksheet with data here.
            }

            // Flag indicating whether to enforce all columns on one page
            bool fitAllColumnsOnOnePage = true;

            if (fitAllColumnsOnOnePage)
            {
                // Configure each worksheet to fit all columns on a single page
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    ws.PageSetup.FitToPagesWide = 1; // one page wide
                    ws.PageSetup.FitToPagesTall = 0; // height adjusts automatically
                }

                // Rendering options used for validation
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
                {
                    AllColumnsInOnePagePerSheet = true,
                    OnePagePerSheet = true
                };

                // Validate that each worksheet fits on one page
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    SheetRender render = new SheetRender(ws, renderOptions);
                    int pageCount = render.PageCount;
                    Console.WriteLine($"Worksheet \"{ws.Name}\" page count: {pageCount}");
                    if (pageCount != 1)
                    {
                        Console.WriteLine("Warning: Columns do not fit within a single page.");
                    }
                }

                // Save the workbook as PDF with the same fitting options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    AllColumnsInOnePagePerSheet = true,
                    OnePagePerSheet = true
                };
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook saved as PDF to \"{outputPath}\".");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
