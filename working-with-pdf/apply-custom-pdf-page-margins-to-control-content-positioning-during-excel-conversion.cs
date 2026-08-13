// Title: Set Custom PDF Page Margins in C# with Aspose.Cells
// Description: Demonstrates how to define top, bottom, left, and right margins (in centimeters) via Worksheet.PageSetup, apply PdfSaveOptions (including OnePagePerSheet), and export an Excel workbook to a PDF with precise margin control.
// Keywords: Aspose.Cells PDF margins | C# set page margins Excel to PDF | Worksheet.PageSetup margin cm | PdfSaveOptions OnePagePerSheet | Aspose.Cells .NET export PDF | custom PDF layout C#
// Common Searches: Aspose.Cells set PDF margins C# | how to change page margins when converting Excel to PDF | PdfSaveOptions margin settings example | OnePagePerSheet Aspose.Cells usage | C# export Excel with custom margins
// Developer Intent: Apply precise page‑margin values to a PDF generated from an Excel workbook.
// Use Cases: Create a printable report with a 2 cm top margin for a header banner. | Generate invoices where left/right margins are 1.5 cm to meet standard paper constraints. | Export a single‑sheet financial summary to one PDF page while preserving custom spacing.
// AI Prompts: Show C# code that sets top, bottom, left, and right margins in centimeters for a PDF created with Aspose.Cells. | Provide an example of using PdfSaveOptions.OnePagePerSheet together with custom margins in Aspose.Cells .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsMarginExample
{
    // Demonstrates how to define top, bottom, left, and right margins (in centimeters) via Worksheet.PageSetup, apply PdfSaveOptions (including OnePagePerSheet), and export an Excel workbook to a PDF with precise margin control.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue("Row 1");
            sheet.Cells["B2"].PutValue(123);
            sheet.Cells["A3"].PutValue("Row 2");
            sheet.Cells["B3"].PutValue(456);

            // Configure custom page margins (values are in centimeters)
            PageSetup pageSetup = sheet.PageSetup;
            pageSetup.TopMargin = 2.0;    // 2 cm top margin
            pageSetup.BottomMargin = 1.0; // 1 cm bottom margin
            pageSetup.LeftMargin = 1.5;   // 1.5 cm left margin
            pageSetup.RightMargin = 1.5;  // 1.5 cm right margin

            // Create PDF save options (optional: set OnePagePerSheet to keep layout)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            pdfOptions.OnePagePerSheet = true;

            // Save the workbook to PDF with the custom margins applied
            workbook.Save("CustomMarginsOutput.pdf", pdfOptions);

            Console.WriteLine("PDF generated with custom page margins.");
        }
    }
}
