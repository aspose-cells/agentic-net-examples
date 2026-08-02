// Title: Apply 600‑point custom paper size, export to PDF, then reset to A4 – Aspose.Cells .NET
// Description: Shows how to set a 600‑point (≈8.33 in) square page size on a worksheet via PageSetup.CustomPaperSize, save the workbook as a PDF, then restore the standard A4 size and save a second PDF.
// Keywords: Aspose.Cells | custom paper size | PageSetup.CustomPaperSize | 600 points | PDF export .NET | reset to A4 | worksheet page setup | points to inches conversion
// Common Searches: Aspose.Cells set custom page size in points | How to export worksheet to PDF with non‑standard dimensions | Reset worksheet paper size to A4 after PDF generation | Convert points to inches for PageSetup in Aspose.Cells | C# example custom paper size PDF
// Developer Intent: Create a PDF with a 600‑point square page, then revert the worksheet to A4 for a second PDF.
// Use Cases: Produce a square‑format PDF brochure while keeping the original workbook for standard A4 printing. | Temporarily enlarge the page to match a large chart, export, and then return to default size for other sections. | Generate two versions of the same workbook—one custom‑sized PDF and one regular A4 PDF.
// AI Prompts: Write C# code that sets a 600‑point custom paper size on a worksheet and saves it as PDF using Aspose.Cells. | Show how to convert 600 points to inches and apply PageSetup.CustomPaperSize, then reset to PaperA4. | Explain the steps to export a workbook with a custom page size and then create an A4 version without recreating the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Shows how to set a 600‑point (≈8.33 in) square page size on a worksheet via PageSetup.CustomPaperSize, save the workbook as a PDF, then restore the standard A4 size and save a second PDF.
class CustomPaperSizeDemo
{
    static void Main()
    {
        // Create a new workbook and obtain the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add some sample content (optional, just for illustration)
        sheet.Cells["A1"].PutValue("Custom Paper Size Demo");
        sheet.Cells["A2"].PutValue("Width and Height = 600 points (≈8.33 inches)");

        // Convert 600 points to inches (1 point = 1/72 inch)
        double inches = 600.0 / 72.0; // ≈8.3333 inches

        // Apply a custom paper size of 600 points × 600 points
        sheet.PageSetup.CustomPaperSize(inches, inches);

        // Render the worksheet to PDF using the custom size
        PdfSaveOptions pdfOptions = new PdfSaveOptions(); // inherits PaginatedSaveOptions
        workbook.Save("CustomSize.pdf", pdfOptions);

        // Revert the worksheet to the standard A4 paper size
        sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;

        // Save again to demonstrate the reverted size
        workbook.Save("A4Size.pdf", SaveFormat.Pdf);
    }
}
