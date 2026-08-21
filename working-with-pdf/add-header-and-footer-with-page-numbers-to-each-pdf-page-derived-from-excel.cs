// Title: Add Header and Footer with Page Numbers to PDF Pages from Excel using Aspose.Cells (C#)
// Description: Creates or loads a Workbook, inserts data, configures each worksheet's PageSetup to show "Page &P of &N" in the center of the header and footer, and saves the result as a PDF with page‑numbered headers and footers.
// Keywords: Aspose.Cells | C# PDF header footer | page numbers Excel to PDF | SetHeader SetFooter Aspose | PdfSaveOptions | Excel pagination
// Common Searches: Aspose.Cells add page numbers to PDF | C# set header footer when converting Excel to PDF | How to show Page X of Y in PDF generated from Excel | Aspose.Cells PDF export header footer example | Add pagination to PDF using Aspose.Cells .NET
// Developer Intent: Insert dynamic page‑number headers and footers into every PDF page produced from an Excel workbook.
// Use Cases: Generating multi‑page PDF reports that display "Page X of Y" in both header and footer. | Automating invoice PDFs where each page must include pagination for legal compliance. | Creating printable Excel‑to‑PDF documents with consistent pagination across all worksheets.
// AI Prompts: Show C# code that uses Aspose.Cells to add a centered "Page &P of &N" header and footer before saving as PDF. | Explain how to customize the header/footer text while keeping dynamic page numbers in Aspose.Cells PDF export. | Provide a step‑by‑step guide to apply the same header/footer settings to all worksheets in a workbook.

using System;
using Aspose.Cells;

// Creates or loads a Workbook, inserts data, configures each worksheet's PageSetup to show "Page &P of &N" in the center of the header and footer, and saves the result as a PDF with page‑numbered headers and footers.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Add some sample data to the first worksheet
        Worksheet firstSheet = workbook.Worksheets[0];
        firstSheet.Cells["A1"].PutValue("Sample Data");

        // Apply header and footer with page numbers to every worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            PageSetup pageSetup = sheet.PageSetup;

            // Center section of header: "Page X of Y"
            pageSetup.SetHeader(1, "Page &P of &N");

            // Center section of footer: "Page X of Y"
            pageSetup.SetFooter(1, "Page &P of &N");
        }

        // Save the workbook as PDF
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        workbook.Save("Output.pdf", pdfOptions);
    }
}
