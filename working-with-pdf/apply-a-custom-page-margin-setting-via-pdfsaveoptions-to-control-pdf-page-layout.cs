// Title: Set Custom PDF Page Margins in C# with Aspose.Cells PdfSaveOptions
// Description: Demonstrates how to define left, right, top, and bottom margins (in centimeters) on a worksheet using PageSetup, configure PdfSaveOptions, and export the workbook to a PDF that respects the specified layout.
// Keywords: Aspose.Cells PDF margins C# | PdfSaveOptions custom margins | PageSetup margin settings | export Excel to PDF with specific margins | C# Aspose.Cells PDF layout control
// Common Searches: Aspose.Cells set PDF page margins | C# change worksheet margins before PDF export | PdfSaveOptions margin configuration example | how to use PageSetup margins with Aspose.Cells | OnePagePerSheet effect on PDF layout
// Developer Intent: Apply precise margin dimensions to a worksheet and generate a PDF that reflects those settings using Aspose.Cells.
// Use Cases: Produce printable PDFs that meet corporate margin standards. | Align Excel‑to‑PDF output with a predefined document template. | Create multi‑sheet PDFs where each sheet retains its own margin configuration.
// AI Prompts: Show how to set page margins in inches instead of centimeters with Aspose.Cells before saving to PDF. | Provide code to assign different margin values to each worksheet and export them as separate PDFs. | Explain the interaction between OnePagePerSheet and custom margins in PdfSaveOptions.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to define left, right, top, and bottom margins (in centimeters) on a worksheet using PageSetup, configure PdfSaveOptions, and export the workbook to a PDF that respects the specified layout.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data
        sheet.Cells["A1"].PutValue("Sample Header");
        sheet.Cells["A2"].PutValue("Row 1");
        sheet.Cells["A3"].PutValue("Row 2");

        // Apply custom page margins (values are in centimeters)
        sheet.PageSetup.LeftMargin = 2.0;    // 2 cm left margin
        sheet.PageSetup.RightMargin = 2.0;   // 2 cm right margin
        sheet.PageSetup.TopMargin = 3.0;     // 3 cm top margin
        sheet.PageSetup.BottomMargin = 3.0;  // 3 cm bottom margin

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Optional: control pagination behavior (example setting)
        pdfOptions.OnePagePerSheet = false;

        // Save the workbook as a PDF file using the custom margin settings
        workbook.Save("CustomMargins.pdf", pdfOptions);
    }
}
