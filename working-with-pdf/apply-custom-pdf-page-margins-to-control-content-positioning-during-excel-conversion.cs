// Title: Set PDF page margins with Aspose.Cells for .NET (C#)
// Description: Create a workbook, define top, bottom, left, and right margins in centimeters via PageSetup, configure PdfSaveOptions (e.g., OnePagePerSheet), and save the worksheet as a PDF with precise layout control.
// Keywords: Aspose.Cells PDF margins | C# set page margins | Excel to PDF custom margins | PageSetup centimeters | PdfSaveOptions OnePagePerSheet | Aspose.Cells .NET PDF export
// Common Searches: Aspose.Cells set PDF margins C# | how to change top bottom margins when converting Excel to PDF | PDF export with custom left right margins Aspose.Cells | OnePagePerSheet option Aspose.Cells PDF | page setup margin units centimeters Aspose.Cells
// Developer Intent: Define exact page margin values before exporting an Excel worksheet to PDF.
// Use Cases: Produce printable PDFs that align with corporate templates by specifying exact centimeter margins. | Generate multi‑sheet reports where each sheet starts on a new PDF page with uniform margins. | Create PDF invoices where content must begin a set distance from the page edges.
// AI Prompts: Show C# code to set top, bottom, left, and right PDF margins in centimeters using Aspose.Cells. | How can I export an Excel workbook to PDF with one page per sheet and custom margins in Aspose.Cells? | Retrieve and modify the current PageSetup margins of a worksheet before saving it as a PDF with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsPdfMarginDemo
{
    // Create a workbook, define top, bottom, left, and right margins in centimeters via PageSetup, configure PdfSaveOptions (e.g., OnePagePerSheet), and save the worksheet as a PDF with precise layout control.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data (optional, just to have visible content)
            sheet.Cells["A1"].PutValue("Demo of custom PDF margins");
            sheet.Cells["A2"].PutValue("Top margin: 1.5 cm");
            sheet.Cells["A3"].PutValue("Bottom margin: 2.0 cm");
            sheet.Cells["A4"].PutValue("Left margin: 1.0 cm");
            sheet.Cells["A5"].PutValue("Right margin: 1.0 cm");

            // Configure page margins (unit: centimeters)
            PageSetup pageSetup = sheet.PageSetup;
            pageSetup.TopMargin = 1.5;    // 1.5 cm top margin
            pageSetup.BottomMargin = 2.0; // 2.0 cm bottom margin
            pageSetup.LeftMargin = 1.0;   // 1.0 cm left margin
            pageSetup.RightMargin = 1.0;  // 1.0 cm right margin

            // Create PDF save options (optional, can customize further)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            // Example: keep each worksheet on a separate page
            pdfOptions.OnePagePerSheet = true;

            // Save the workbook as PDF with the defined margins (lifecycle: save)
            workbook.Save("CustomMarginsOutput.pdf", pdfOptions);

            Console.WriteLine("Workbook saved to PDF with custom margins.");
        }
    }
}
