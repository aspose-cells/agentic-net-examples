// Title: C# – Preserve Clickable Hyperlinks When Converting Excel to PDF with Aspose.Cells
// Description: Demonstrates how to add a hyperlink to a cell, set custom display text, and save the workbook as a PDF. Aspose.Cells automatically retains the link, so the PDF contains a functional, clickable URL.
// Keywords: Aspose.Cells PDF hyperlink | C# Excel to PDF conversion | clickable links in PDF | preserve hyperlinks Aspose | save workbook as PDF C# | Aspose.Cells hyperlink support
// Common Searches: keep hyperlinks when exporting Excel to PDF Aspose.Cells | C# Aspose.Cells preserve links in PDF output | how to make PDF links clickable from Excel | Aspose.Cells PDF export hyperlink retention
// Developer Intent: Ensure that URLs embedded in Excel cells remain active after exporting the workbook to PDF using Aspose.Cells for .NET.
// Use Cases: Generating PDF reports that include reference links for end‑users. | Creating marketing PDFs from spreadsheets where product pages must stay reachable. | Automating batch export of data sheets while preserving navigation through embedded URLs.
// AI Prompts: Write C# code with Aspose.Cells to add several hyperlinks to different cells and verify they are clickable in the exported PDF. | Explain the internal process Aspose.Cells uses to retain hyperlinks during PDF conversion and list any optional settings that influence this behavior. | Show how to set a hyperlink’s display text different from its URL and test the link in the resulting PDF.

using Aspose.Cells;
using System;

// Demonstrates how to add a hyperlink to a cell, set custom display text, and save the workbook as a PDF. Aspose.Cells automatically retains the link, so the PDF contains a functional, clickable URL.
class PreserveHyperlinkPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Set the display text for the cell
        sheet.Cells["A1"].PutValue("Visit Aspose");

        // Add a hyperlink to the cell A1
        // Parameters: start cell, total rows, total columns, hyperlink address
        sheet.Hyperlinks.Add("A1", 1, 1, "https://www.aspose.com");

        // Save the workbook as PDF.
        // Hyperlinks are preserved automatically when saving to PDF.
        workbook.Save("HyperlinkPreserved.pdf", SaveFormat.Pdf);
    }
}
