// Title: Aspose.Cells for .NET – Place a custom footer at (0,0) on every PDF page (C#)
// Description: Create a workbook, add optional data, set the left‑section footer text, set FooterMargin to 0, and save as PDF so the footer appears at the bottom‑left corner (coordinate 0,0) on each page.
// Keywords: Aspose.Cells | C# | PDF footer | footer margin zero | custom footer text | set footer position | page setup footer | Aspose.Cells PDF export | footer coordinates | Aspose.Cells example
// Common Searches: Aspose.Cells set footer at bottom of PDF | C# footer margin zero Aspose.Cells PDF | place custom footer (0,0) in generated PDF using Aspose.Cells
// Developer Intent: Add a left‑section footer with custom text and position it at the page origin (0,0) for every page of a PDF generated from a workbook.
// Use Cases: Add a legal disclaimer that must sit flush with the bottom edge of each PDF page. | Create invoices where terms and conditions appear as a footer anchored to the page margin. | Generate multi‑page reports that require a consistent footer aligned to the bottom‑left corner.
// AI Prompts: Show C# code using Aspose.Cells to set a custom footer at coordinate (0,0) for all PDF pages. | How do I configure FooterMargin = 0 and add left‑section text to position a footer at the bottom edge of a PDF with Aspose.Cells? | Provide an Aspose.Cells example that places a custom footer at the page origin when saving a workbook as PDF.

using System;
using Aspose.Cells;

namespace AsposeCellsFooterExample
{
    // Create a workbook, add optional data, set the left‑section footer text, set FooterMargin to 0, and save as PDF so the footer appears at the bottom‑left corner (coordinate 0,0) on each page.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add some sample data (optional, just to have content)
            worksheet.Cells["A1"].PutValue("Sample content for PDF conversion.");

            // Configure the footer:
            // Set the left section (index 0) to the desired custom text.
            // Setting FooterMargin to 0 positions the footer at the very bottom edge.
            // This effectively places the text at coordinate (0,0) relative to the page margin.
            worksheet.PageSetup.SetFooter(0, "Custom Footer Text");
            worksheet.PageSetup.FooterMargin = 0; // centimeters from bottom edge

            // Save the workbook as PDF
            workbook.Save("CustomFooter.pdf", SaveFormat.Pdf);
        }
    }
}
