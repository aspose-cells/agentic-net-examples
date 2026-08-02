// Title: Aspose.Cells .NET – Create PDF with Custom 4×6‑inch Paper Size and 300 DPI Image Quality
// Description: Demonstrates how to set a 4 × 6 in custom paper size, configure worksheet PrintQuality to 300 DPI, and use PdfSaveOptions.SetImageResample(300, 90) for high‑resolution JPEG output, then save the workbook as a PDF with crisp graphics.
// Keywords: Aspose.Cells | .NET | C# | PDF export | custom paper size | 4x6 inches | 300 DPI | PrintQuality | PdfSaveOptions | SetImageResample | high resolution | JPEG quality | worksheet page setup
// Common Searches: Aspose.Cells set custom paper size PDF | How to export PDF at 300 DPI using Aspose.Cells | PdfSaveOptions image resample example C# | Increase PDF image quality Aspose.Cells .NET | PrintQuality DPI worksheet Aspose.Cells
// Developer Intent: Export a worksheet to a PDF with exact 4 × 6 in dimensions and 300 DPI image resolution.
// Use Cases: Generate small‑format flyers (4×6 in) with sharp images for marketing materials. | Create high‑resolution invoices or receipts where barcodes and logos must be clear. | Produce mobile‑friendly sales reports that retain image fidelity on handheld printers.
// AI Prompts: Show how to change the custom paper size to 5×7 in while keeping 300 DPI image quality in Aspose.Cells. | Explain the impact of JPEG quality values in SetImageResample on file size and visual fidelity. | Provide code to apply different DPI settings to multiple worksheets and save each as a separate PDF.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfDpiDemo
{
    // Demonstrates how to set a 4 × 6 in custom paper size, configure worksheet PrintQuality to 300 DPI, and use PdfSaveOptions.SetImageResample(300, 90) for high‑resolution JPEG output, then save the workbook as a PDF with crisp graphics.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(150);
            sheet.Cells["A3"].PutValue("Bananas");
            sheet.Cells["B3"].PutValue(200);

            // ------------------------------------------------------------
            // 1. Set a custom paper size (width: 4 inches, height: 6 inches)
            // ------------------------------------------------------------
            sheet.PageSetup.CustomPaperSize(4.0, 6.0);

            // ------------------------------------------------------------
            // 2. Increase the print quality (DPI) for the worksheet
            // ------------------------------------------------------------
            sheet.PageSetup.PrintQuality = 300; // 300 DPI

            // ------------------------------------------------------------
            // 3. Configure PDF save options to resample images at 300 DPI
            //    and use a high JPEG quality (e.g., 90%)
            // ------------------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            pdfOptions.SetImageResample(300, 90); // desiredPPI = 300, jpegQuality = 90

            // ------------------------------------------------------------
            // 4. Save the workbook as PDF with the above settings
            // ------------------------------------------------------------
            workbook.Save("CustomPaper_300DPI.pdf", pdfOptions);

            Console.WriteLine("PDF generated with custom paper size and 300 DPI image quality.");
        }
    }
}
