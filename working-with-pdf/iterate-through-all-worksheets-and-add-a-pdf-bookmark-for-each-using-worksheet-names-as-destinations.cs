// Title: C# – Add PDF Bookmarks for Every Worksheet with Aspose.Cells
// Description: Shows how to export a workbook to PDF and automatically generate top‑level bookmarks that link each worksheet’s name to its A1 cell, using PdfBookmarkEntry and PdfSaveOptions in Aspose.Cells for .NET.
// Keywords: Aspose.Cells PDF bookmark | C# PDF export Aspose.Cells | PdfBookmarkEntry example | worksheet to PDF navigation | Aspose.Cells .NET bookmark tutorial
// Common Searches: Aspose.Cells create PDF bookmarks for worksheets | C# export workbook to PDF with bookmarks | set PDF bookmark destination cell Aspose | generate PDF with sheet name bookmarks .NET
// Developer Intent: Produce a PDF where each worksheet is represented by a clickable bookmark that jumps to cell A1 of that sheet.
// Use Cases: Deliver a multi‑sheet financial report where users can jump directly to Finance, HR, or IT sections via PDF bookmarks. | Provide quick navigation in exported PDFs of large workbooks for end‑users. | Automate bookmark creation when worksheets are added dynamically at runtime.
// AI Prompts: Write C# code that uses Aspose.Cells to add a PDF bookmark for every worksheet, using the sheet name as the title and cell A1 as the target. | Explain how to build a root PdfBookmarkEntry and attach sub‑entries for each worksheet before saving to PDF. | Give an example of saving a workbook with PdfSaveOptions that includes a custom bookmark hierarchy.

using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfBookmarks
{
    // Shows how to export a workbook to PDF and automatically generate top‑level bookmarks that link each worksheet’s name to its A1 cell, using PdfBookmarkEntry and PdfSaveOptions in Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Example: add some worksheets with sample data
            Worksheet sheet1 = workbook.Worksheets.Add("Finance");
            sheet1.Cells["A1"].PutValue("Finance Overview");

            Worksheet sheet2 = workbook.Worksheets.Add("HR");
            sheet2.Cells["A1"].PutValue("HR Overview");

            Worksheet sheet3 = workbook.Worksheets.Add("IT");
            sheet3.Cells["A1"].PutValue("IT Overview");

            // Create a root bookmark entry with no text.
            // Children of this entry will appear as top‑level bookmarks.
            PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
            {
                Text = null,
                SubEntry = new ArrayList()
            };

            // Iterate through all worksheets and create a bookmark for each.
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Ensure the destination cell exists (A1 is used here).
                // You can choose any cell that makes sense as the bookmark target.
                PdfBookmarkEntry entry = new PdfBookmarkEntry
                {
                    Text = ws.Name,          // Bookmark title = worksheet name
                    Destination = ws.Cells["A1"] // Destination cell
                };

                // Add the entry to the root's sub‑entries.
                rootBookmark.SubEntry.Add(entry);
            }

            // Configure PDF save options with the constructed bookmark hierarchy.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Bookmark = rootBookmark
            };

            // Save the workbook as a PDF with the bookmarks.
            workbook.Save("WorkbookWithBookmarks.pdf", pdfOptions);
        }
    }
}
