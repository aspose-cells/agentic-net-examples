// Title: Export Workbook to PDF with One Page Per Worksheet using Aspose.Cells (C#)
// Description: Creates a workbook with multiple worksheets, enables PdfSaveOptions.OnePagePerSheet, and saves the workbook as a single PDF where each worksheet begins on a new page.
// Keywords: Aspose.Cells | PdfSaveOptions | OnePagePerSheet | C# PDF export | Excel to PDF per sheet | save workbook as PDF | separate pages per worksheet | Aspose.Cells PDF options
// Common Searches: Aspose.Cells OnePagePerSheet C# example | export Excel worksheets to PDF each on separate page | PdfSaveOptions OnePagePerSheet usage | C# save workbook as PDF with one page per sheet | how to generate PDF from multiple worksheets using Aspose.Cells
// Developer Intent: Generate a PDF where every worksheet is rendered on its own page by setting PdfSaveOptions.OnePagePerSheet to true.
// Use Cases: Produce a multi‑sheet report PDF with each sheet starting on a new page. | Automate invoice generation where each worksheet represents an invoice and appears as a distinct PDF page. | Create a PDF handbook from a workbook, using one worksheet per chapter for clear navigation.
// AI Prompts: Write C# code that exports an Aspose.Cells workbook to PDF with OnePagePerSheet enabled and sets image quality to high. | Show how to save the PDF to a MemoryStream instead of a file while keeping OnePagePerSheet true. | Explain how to combine PdfSaveOptions.OnePagePerSheet with password protection and other PDF settings in Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a workbook with multiple worksheets, enables PdfSaveOptions.OnePagePerSheet, and saves the workbook as a single PDF where each worksheet begins on a new page.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Populate the first worksheet with sample data
        Worksheet sheet1 = workbook.Worksheets[0];
        for (int i = 0; i < 10; i++)
        {
            sheet1.Cells[i, 0].Value = $"Sheet1 Data {i + 1}";
        }

        // Add a second worksheet and populate it with sample data
        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
        for (int i = 0; i < 10; i++)
        {
            sheet2.Cells[i, 0].Value = $"Sheet2 Data {i + 1}";
        }

        // Configure PDF save options to output each sheet on a separate page
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.OnePagePerSheet = true;

        // Save the workbook as a PDF file using the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}
