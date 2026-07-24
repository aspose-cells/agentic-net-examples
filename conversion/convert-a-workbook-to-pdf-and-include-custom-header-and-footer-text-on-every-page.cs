// Title: Add Header & Footer to Excel and Export as PDF with Aspose.Cells (C#)
// Description: Demonstrates how to create a workbook, define left/center/right header and footer sections using PageSetup.SetHeader/SetFooter, populate sample rows, configure PdfSaveOptions, and save the file as a PDF where the custom header/footer appear on every page.
// Keywords: Aspose.Cells C# header footer PDF | Excel to PDF custom header | SetHeader SetFooter Aspose.Cells | PdfSaveOptions export PDF | page number and date in PDF | C# convert workbook to PDF | Aspose.Cells printable report | Excel pagination header footer
// Common Searches: how to add header and footer to Excel PDF using Aspose.Cells | Aspose.Cells C# set left center right header before PDF export | include page number date file name in PDF generated from Excel | Aspose.Cells PDF export with confidential footer text | C# Aspose.Cells example for custom PDF headers
// Developer Intent: Generate a PDF from an Excel workbook that consistently displays custom header and footer text on each page.
// Use Cases: Brand‑consistent printable reports with company logo and page numbers. | Multi‑page invoices that show page X of Y, date, and file name. | Automated export of large data sheets where each PDF page includes a confidentiality notice and timestamp.
// AI Prompts: Provide C# code using Aspose.Cells to insert an image logo in the left header and export to PDF. | Show how to hide the footer on the first page while keeping it on subsequent pages during PDF conversion. | Explain how to enable PDF/A compliance and embed fonts in PdfSaveOptions while preserving custom headers and footers.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsHeaderFooterPdfDemo
{
    // Demonstrates how to create a workbook, define left/center/right header and footer sections using PageSetup.SetHeader/SetFooter, populate sample rows, configure PdfSaveOptions, and save the file as a PDF where the custom header/footer appear on every page.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (you can repeat for other worksheets if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Set custom header (left, center, right sections)
            // -------------------------------------------------
            // Left section: custom text with font and size
            sheet.PageSetup.SetHeader(0, "&\"Arial,Bold\"&12 My Custom Header");
            // Center section: page number and total pages
            sheet.PageSetup.SetHeader(1, "Page &P of &N");
            // Right section: current date
            sheet.PageSetup.SetHeader(2, "&D");

            // -------------------------------------------------
            // Set custom footer (left, center, right sections)
            // -------------------------------------------------
            // Left section: file name without path
            sheet.PageSetup.SetFooter(0, "&F");
            // Center section: static text
            sheet.PageSetup.SetFooter(1, "Confidential");
            // Right section: current time
            sheet.PageSetup.SetFooter(2, "&T");

            // Add some sample data to demonstrate pagination
            for (int row = 0; row < 200; row++)
            {
                sheet.Cells[row, 0].PutValue($"Row {row + 1}");
                sheet.Cells[row, 1].PutValue($"Data {row + 1}");
            }

            // -------------------------------------------------
            // Prepare PDF save options (optional customizations)
            // -------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            // Example: keep document structure (can be omitted if not required)
            pdfOptions.ExportDocumentStructure = true;

            // -------------------------------------------------
            // Save the workbook as PDF; headers/footers will appear on every page
            // -------------------------------------------------
            string outputPath = "Workbook_With_HeaderFooter.pdf";
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully saved to PDF with custom header/footer at: {outputPath}");
        }
    }
}
