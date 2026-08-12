// Title: C# – Set 0.5‑inch margins on all sides when exporting Excel to PDF with Aspose.Cells
// Description: Demonstrates how to configure PageSetup.LeftMarginInch, RightMarginInch, TopMarginInch, and BottomMarginInch to 0.5 inches, then save the workbook as a PDF using PdfSaveOptions, producing a document with uniform half‑inch margins.
// Keywords: Aspose.Cells PDF margins | C# set PDF page margins | PdfSaveOptions margin settings | PageSetup margin inches | .NET export Excel to PDF | custom PDF margins Aspose
// Common Searches: Aspose.Cells set 0.5 inch PDF margins C# | export Excel to PDF with specific margins .NET | how to change PDF page margins using Aspose.Cells | C# PdfSaveOptions custom margins example | adjust left right top bottom margins before PDF export
// Developer Intent: Apply a uniform half‑inch margin to every side of a PDF generated from an Excel workbook.
// Use Cases: Create printable reports that must fit within a 0.5‑inch printable area. | Generate invoices or statements with consistent margin layout for branding. | Produce regulatory or legal PDFs that require precise margin dimensions.
// AI Prompts: Write C# code with Aspose.Cells to set 0.5‑inch margins on all sides and save as PDF. | Explain the relationship between PageSetup margin properties and PdfSaveOptions during PDF export. | Show how to specify margins in points or centimeters instead of inches for Aspose.Cells PDF output.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfMarginExample
{
    // Demonstrates how to configure PageSetup.LeftMarginInch, RightMarginInch, TopMarginInch, and BottomMarginInch to 0.5 inches, then save the workbook as a PDF using PdfSaveOptions, producing a document with uniform half‑inch margins.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data so the PDF has visible content
            sheet.Cells["A1"].PutValue("Demo of 0.5 inch margins on all sides");
            sheet.Cells["A2"].PutValue("This text should appear within the defined margins.");

            // Set all page margins to 0.5 inches using the Inch properties
            PageSetup pageSetup = sheet.PageSetup;
            pageSetup.LeftMarginInch = 0.5;
            pageSetup.RightMarginInch = 0.5;
            pageSetup.TopMarginInch = 0.5;
            pageSetup.BottomMarginInch = 0.5;

            // Create PDF save options (no special options needed for margins)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF with the specified margins
            workbook.Save("CustomMargins.pdf", pdfOptions);

            Console.WriteLine("PDF saved with 0.5 inch margins on all sides.");
        }
    }
}
