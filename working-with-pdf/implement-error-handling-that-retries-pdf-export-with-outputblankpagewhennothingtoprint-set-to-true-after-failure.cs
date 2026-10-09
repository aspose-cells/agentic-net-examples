// Title: Add retry logic for PDF export in Aspose.Cells C# by enabling OutputBlankPageWhenNothingToPrint after a failure
// AI Prompts: Generate C# code that saves an Aspose.Cells workbook to PDF, catches any exception, sets PdfSaveOptions.OutputBlankPageWhenNothingToPrint = true, and retries the save operation. | Write a C# method that attempts to export a workbook to PDF with Aspose.Cells, and on failure automatically re‑exports using the blank‑page fallback option.
// Common Searches: Aspose.Cells C# handle exception on PDF conversion and perform a second save with blank page option | Enable OutputBlankPageWhenNothingToPrint on second PDF export attempt using Aspose.Cells | C# catch workbook.Save error and perform fallback export to PDF with blank page in Aspose.Cells
// Tags: Aspose.Cells PDF export with OutputBlankPageWhenNothingToPrint | C# exception handling for workbook.Save PDF conversion | PdfSaveOptions OutputBlankPageWhenNothingToPrint | Aspose.Cells PDF conversion error handling

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads or creates a workbook, tries to save it as PDF, and if the first save fails, sets OutputBlankPageWhenNothingToPrint to true and retries, logging any errors.
class Program
{
    static void Main()
    {
        // Create or load a workbook.
        Workbook workbook;
        string templatePath = "template.xlsx";

        try
        {
            if (File.Exists(templatePath))
            {
                // Load existing workbook if the template file is present.
                workbook = new Workbook(templatePath);
            }
            else
            {
                // Create a new workbook with a default worksheet.
                workbook = new Workbook();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load or create workbook: {ex.Message}");
            return;
        }

        // Prepare PDF save options with default settings.
        var pdfOptions = new PdfSaveOptions();

        try
        {
            // Attempt to export PDF with the initial options.
            workbook.Save("output.pdf", pdfOptions);
        }
        catch (Exception)
        {
            // If export fails, enable OutputBlankPageWhenNothingToPrint and retry.
            pdfOptions.OutputBlankPageWhenNothingToPrint = true;

            try
            {
                workbook.Save("output.pdf", pdfOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"PDF export failed: {ex.Message}");
            }
        }
    }
}
