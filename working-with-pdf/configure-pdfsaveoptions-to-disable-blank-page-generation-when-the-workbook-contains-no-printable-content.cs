// Title: C# – Disable Blank Page When Saving an Empty Workbook to PDF with Aspose.Cells
// Description: Shows how to stop Aspose.Cells from inserting a blank page during PDF export when the workbook contains no printable data by clearing the print area and setting PdfSaveOptions.OutputBlankPageWhenNothingToPrint to false.
// Keywords: Aspose.Cells PDF export | PdfSaveOptions | OutputBlankPageWhenNothingToPrint | disable blank page C# | empty workbook to PDF | print area clear Aspose.Cells
// Common Searches: Aspose.Cells prevent blank page PDF export | PdfSaveOptions OutputBlankPageWhenNothingToPrint false example | C# save empty workbook as PDF without blank page | how to hide blank pages in Aspose.Cells PDF conversion
// Developer Intent: Avoid generating a blank PDF page when the workbook has no printable content.
// Use Cases: Create PDF reports that may have zero rows and need to omit the first empty page. | Export invoices where optional sections are hidden, ensuring the final PDF has no extra pages. | Batch‑convert workbooks, some of which are empty, while suppressing unwanted blank pages.
// AI Prompts: Provide C# code using Aspose.Cells to export a workbook to PDF without a blank page when the print area is empty. | Explain the purpose of PdfSaveOptions.OutputBlankPageWhenNothingToPrint and when it should be set to false.

using System;
using Aspose.Cells;

// Shows how to stop Aspose.Cells from inserting a blank page during PDF export when the workbook contains no printable data by clearing the print area and setting PdfSaveOptions.OutputBlankPageWhenNothingToPrint to false.
class DisableBlankPageDemo
{
    static void Main()
    {
        try
        {
            // Create a new empty workbook
            Workbook workbook = new Workbook();

            // Get the default worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Keep the worksheet visible but exclude it from printing by clearing the print area
            sheet.PageSetup.PrintArea = string.Empty;

            // Configure PDF save options: do NOT output a blank page when nothing is printable
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OutputBlankPageWhenNothingToPrint = false
            };

            // Save the workbook to PDF using the configured options
            string outputPath = "NoBlankPage.pdf";
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
