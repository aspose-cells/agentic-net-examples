// Title: C# – Apply a 600‑point custom paper size in Aspose.Cells, export to PDF, then revert to A4
// Description: Demonstrates how to convert 600 points to inches, set a square custom paper size on a worksheet using PageSetup.CustomPaperSize, save the workbook as a PDF, switch the PaperSize back to A4, and generate a second PDF with standard dimensions.
// Keywords: Aspose.Cells custom paper size | C# PageSetup.CustomPaperSize | set worksheet size points | export worksheet to PDF | revert to A4 Aspose.Cells | 600 points to inches conversion
// Common Searches: Aspose.Cells set custom paper size C# | how to export PDF with custom dimensions Aspose.Cells | change worksheet paper size back to A4 after PDF export | convert points to inches for Aspose.Cells PageSetup
// Developer Intent: Configure a worksheet with a 600‑point (≈8.33 in) custom page size, generate a PDF, then restore the default A4 size and generate another PDF.
// Use Cases: Create a square‑format PDF for a brochure while keeping the workbook’s default layout for other outputs. | Produce two PDFs from the same data: one for custom‑size printing, another for standard A4 distribution. | Temporarily adjust page dimensions for a single export operation without altering the workbook’s permanent settings.
// AI Prompts: Generate C# code that sets a worksheet's page size to 600 points, saves a PDF, then resets to A4 and saves a second PDF using Aspose.Cells. | Explain why PaperSize must be set to Custom before calling CustomPaperSize and how to convert points to inches for Aspose.Cells. | Provide a reusable C# method that accepts a point value, applies it as a custom paper size, exports to PDF, and restores the original paper size.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsCustomPaperDemo
{
    // Demonstrates how to convert 600 points to inches, set a square custom paper size on a worksheet using PageSetup.CustomPaperSize, save the workbook as a PDF, switch the PaperSize back to A4, and generate a second PDF with standard dimensions.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Convert 600 points to inches (1 point = 1/72 inch)
            double inches = 600.0 / 72.0; // ≈ 8.3333 inches

            // Apply a custom paper size of 600 points (width x height)
            sheet.PageSetup.CustomPaperSize(inches, inches);
            // Indicate that a custom size is being used
            sheet.PageSetup.PaperSize = PaperSizeType.Custom;

            // Render the worksheet to PDF with the custom paper size
            workbook.Save("CustomPaperSize.pdf", SaveFormat.Pdf);

            // Revert the paper size back to standard A4
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;

            // Render the worksheet to PDF again, now using A4 size
            workbook.Save("A4PaperSize.pdf", SaveFormat.Pdf);
        }
    }
}
