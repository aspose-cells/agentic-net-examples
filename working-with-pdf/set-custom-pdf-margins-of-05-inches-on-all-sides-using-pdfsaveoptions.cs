// Title: Aspose.Cells .NET – Apply 0.5‑inch margins to PDF export
// Description: Create a workbook, configure each side to 0.5 in using PageSetup, then export to PDF with PdfSaveOptions for consistent borders.
// Keywords: Aspose.Cells PDF margins | C# PdfSaveOptions | PageSetup margin inches | half‑inch PDF border | Aspose.Cells export settings
// Common Searches: set pdf margins aspose.cells c# | half inch margins pdf export .net | page setup margin inches aspose | pdfsaveoptions margin configuration | aspocells pdf margin example
// Developer Intent: Configure a workbook so the generated PDF has 0.5 in margins on every side.
// Use Cases: Print‑ready reports requiring a half‑inch printable area | Invoices that must meet corporate margin standards | Compliance documents where exact margin dimensions are mandated
// AI Prompts: How do I set 0.5 inch margins for a PDF generated with Aspose.Cells in C#? | Provide a C# example that uses PageSetup and PdfSaveOptions to export a workbook with uniform half‑inch margins. | Explain margin adjustments for different page sizes when exporting to PDF with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsMarginExample
{
    // Create a workbook, configure each side to 0.5 in using PageSetup, then export to PDF with PdfSaveOptions for consistent borders.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data (optional, just to see the effect)
            sheet.Cells["A1"].PutValue("Demo of 0.5 inch margins on all sides");

            // Set all page margins to 0.5 inches using the Inch properties
            PageSetup pageSetup = sheet.PageSetup;
            pageSetup.LeftMarginInch = 0.5;
            pageSetup.RightMarginInch = 0.5;
            pageSetup.TopMarginInch = 0.5;
            pageSetup.BottomMarginInch = 0.5;

            // Create PDF save options (can customize further if needed)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF with the specified margins
            workbook.Save("MarginsHalfInch.pdf", pdfOptions);

            Console.WriteLine("PDF saved with 0.5 inch margins on all sides.");
        }
    }
}
