// Title: Add Nested PDF Bookmarks (Parent‑Child‑Sub‑child) with Aspose.Cells for .NET
// Description: Demonstrates how to create a PDF bookmark hierarchy in C# by building a root PdfBookmarkEntry, adding child and sub‑child entries, linking each to a worksheet cell, and exporting the workbook to PDF using PdfSaveOptions.
// Keywords: Aspose.Cells PDF bookmarks | PdfBookmarkEntry hierarchy | nested PDF bookmarks C# | export Excel to PDF with bookmarks | PdfSaveOptions Bookmark property | Aspose.Cells .NET example | Excel worksheet PDF outline
// Common Searches: Aspose.Cells create parent child PDF bookmarks | C# add sub‑bookmarks to PDF export | PdfBookmarkEntry example Aspose.Cells | how to build PDF bookmark tree from Excel | export workbook to PDF with outline
// Developer Intent: Create a parent PDF bookmark with child and sub‑child entries, link each to specific cells, and generate a PDF that displays the full bookmark outline.
// Use Cases: Generate a navigable PDF report where each worksheet appears as a top‑level bookmark. | Produce a PDF document with chapters and sub‑chapters linked to Excel cells for quick reference. | Provide an expandable outline view in PDFs created from Excel workbooks for end‑user navigation.
// AI Prompts: Show me C# code to build a multi‑level PdfBookmarkEntry hierarchy and attach it to PdfSaveOptions in Aspose.Cells. | Explain how to associate each PdfBookmarkEntry with a cell range and keep the root bookmark expanded on PDF open. | Give an example of adding additional sub‑bookmarks under an existing child bookmark in a PDF generated from a workbook.

using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfBookmarkDemo
{
    // Demonstrates how to create a PDF bookmark hierarchy in C# by building a root PdfBookmarkEntry, adding child and sub‑child entries, linking each to a worksheet cell, and exporting the workbook to PDF using PdfSaveOptions.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add three worksheets
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0]; // default sheet
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");

            // Put sample values that will serve as bookmark destinations
            sheet1.Cells["A1"].PutValue("Content of Sheet1");
            sheet2.Cells["A1"].PutValue("Content of Sheet2");
            sheet3.Cells["A1"].PutValue("Content of Sheet3");

            // ---------- Create PDF bookmark hierarchy ----------
            // Root bookmark (will appear at top level)
            PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
            {
                Text = "Workbook Overview",
                Destination = sheet1.Cells["A1"],
                IsOpen = true,
                SubEntry = new ArrayList()
            };

            // First child bookmark
            PdfBookmarkEntry child1 = new PdfBookmarkEntry
            {
                Text = "Section - Sheet2",
                Destination = sheet2.Cells["A1"]
            };

            // Second child bookmark with its own sub‑bookmark
            PdfBookmarkEntry child2 = new PdfBookmarkEntry
            {
                Text = "Section - Sheet3",
                Destination = sheet3.Cells["A1"],
                SubEntry = new ArrayList()
            };

            // Sub‑bookmark under child2
            PdfBookmarkEntry subChild = new PdfBookmarkEntry
            {
                Text = "Sub‑section in Sheet3",
                Destination = sheet3.Cells["A1"] // could point elsewhere if needed
            };

            // Build the hierarchy
            child2.SubEntry.Add(subChild);
            rootBookmark.SubEntry.Add(child1);
            rootBookmark.SubEntry.Add(child2);

            // ---------- Configure PDF save options ----------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Attach the bookmark hierarchy to the PDF
                Bookmark = rootBookmark,
                // Ensure the document structure is exported (optional but common)
                ExportDocumentStructure = true
            };

            // Save the workbook as a PDF with the defined bookmarks
            workbook.Save("WorkbookWithBookmarks.pdf", pdfOptions);

            Console.WriteLine("PDF saved successfully with bookmark hierarchy.");
        }
    }
}
