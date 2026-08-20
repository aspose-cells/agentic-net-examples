// Title: Add Workbook Name as Centered Header on Every PDF Page with Aspose.Cells for .NET
// Description: Demonstrates how to set the workbook name as a centered page header on each worksheet, then export the workbook to PDF using PdfSaveOptions so the header appears on every PDF page.
// Keywords: Aspose.Cells PDF header | C# add workbook name to PDF header | centered page header Aspose.Cells | PdfSaveOptions header example | export workbook to PDF with header | Aspose.Cells .NET PDF export
// Common Searches: Aspose.Cells set header for PDF export | C# add workbook title to PDF page header | centered header on each PDF page Aspose.Cells | how to use PdfSaveOptions for headers | apply same header to all worksheets PDF
// Developer Intent: Insert the workbook's title as a centered header on every page of the PDF generated from an Aspose.Cells workbook.
// Use Cases: Create branded PDF reports that display the workbook title on each page. | Export multiple worksheets to a single PDF while maintaining a consistent header. | Produce printable documents that include the workbook name for easy identification.
// AI Prompts: Generate C# code with Aspose.Cells that adds a left‑aligned date header to each PDF page. | Show how to configure different headers for odd and even pages when saving a workbook to PDF. | Provide an example that adds page numbers to the footer while keeping a static workbook name header.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to set the workbook name as a centered page header on each worksheet, then export the workbook to PDF using PdfSaveOptions so the header appears on every PDF page.
class AddWorkbookNameHeaderToPdf
{
    static void Main()
    {
        // Define a name for the workbook (used in the header)
        string workbookName = "SampleWorkbook";

        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Set a page header that displays the workbook name on the center section of each page
        // Section 1 = Center. The header script can contain plain text.
        sheet.PageSetup.SetHeader(1, workbookName);

        // (Optional) Add some data so the PDF has visible content
        sheet.Cells["A1"].PutValue("Data on first sheet");
        sheet.Cells["A2"].PutValue("More data...");

        // If there are additional worksheets, apply the same header to them
        for (int i = 1; i < workbook.Worksheets.Count; i++)
        {
            Worksheet ws = workbook.Worksheets[i];
            ws.PageSetup.SetHeader(1, workbookName);
        }

        // Configure PDF save options (no special options needed for the header)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the workbook as a PDF; the header will appear on every page
        workbook.Save("WorkbookWithHeader.pdf", pdfOptions);
    }
}
