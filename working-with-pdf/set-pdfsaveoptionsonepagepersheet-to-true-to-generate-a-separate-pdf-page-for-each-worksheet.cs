// Title: Save each Excel worksheet as a separate PDF page with Aspose.Cells for .NET
// Description: Creates a workbook with two sheets, fills them with data, sets PdfSaveOptions.OnePagePerSheet to true, and saves the file so every worksheet appears on its own PDF page.
// Keywords: Aspose.Cells | C# | .NET | PdfSaveOptions | OnePagePerSheet | export Excel to PDF | multiple sheets PDF | single PDF per sheet | worksheet to PDF page
// Common Searches: Aspose.Cells one page per sheet C# | export each Excel sheet to separate PDF page | PdfSaveOptions OnePagePerSheet example | save workbook as PDF with one page per worksheet | C# generate PDF from multi‑sheet Excel
// Developer Intent: Enable the OnePagePerSheet flag so that saving a workbook produces one PDF page for each worksheet.
// Use Cases: Produce a printable report where each worksheet is a distinct PDF page. | Create a PDF portfolio from a financial model that separates sections by sheet. | Distribute multi‑sheet analysis as a single PDF while preserving sheet boundaries.
// AI Prompts: Show how to set custom page margins while keeping OnePagePerSheet true. | Provide code to write the PDF to a MemoryStream instead of a file with OnePagePerSheet enabled. | Explain how to combine OnePagePerSheet with PDF/A compliance and image quality settings.

using System;
using Aspose.Cells;

// Creates a workbook with two sheets, fills them with data, sets PdfSaveOptions.OnePagePerSheet to true, and saves the file so every worksheet appears on its own PDF page.
class Program
{
    static void Main()
    {
        // Create a new workbook and add sample data to two worksheets
        Workbook workbook = new Workbook();

        // First worksheet (default)
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Sheet1";
        for (int i = 0; i < 10; i++)
        {
            sheet1.Cells[i, 0].PutValue($"Sheet1 Data {i + 1}");
        }

        // Second worksheet
        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
        for (int i = 0; i < 10; i++)
        {
            sheet2.Cells[i, 0].PutValue($"Sheet2 Data {i + 1}");
        }

        // Create PDF save options and set OnePagePerSheet to true
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.OnePagePerSheet = true; // Each worksheet will be rendered on a separate PDF page

        // Save the workbook as a PDF using the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}
