// Title: Export Excel to PDF with hierarchical bookmarks using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook with three worksheets, assign each sheet's A1 cell as a bookmark destination, build a root PdfBookmarkEntry with two child entries, enable ExportDocumentStructure, attach the hierarchy to PdfSaveOptions, and save the workbook as a PDF that shows a clickable outline in any PDF viewer.
// Keywords: Aspose.Cells | .NET | PDF bookmarks | ExportDocumentStructure | PdfSaveOptions | Excel to PDF | bookmark hierarchy | outline view
// Common Searches: add bookmarks to PDF generated from Excel Aspose.Cells | export workbook outline with bookmarks Aspose .NET | PdfSaveOptions ExportDocumentStructure true example | create hierarchical PDF bookmarks from multiple worksheets | verify PDF bookmarks appear in outline viewer programmatically
// Developer Intent: Add a multi‑level bookmark outline to a PDF produced from an Excel workbook.
// Use Cases: Produce a PDF report where each worksheet is a top‑level bookmark for fast navigation. | Generate a user manual PDF with a root entry and sub‑bookmarks that jump to specific sections in different sheets. | Create a PDF export that opens on the first sheet while providing a collapsible outline for additional sheets.
// AI Prompts: Show how to add a third‑level nested bookmark to the PDF using Aspose.Cells. | Provide code to programmatically confirm that the bookmarks appear in the PDF outline viewer after saving. | Explain how to set a custom zoom level for each bookmark destination in the generated PDF.

using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfBookmarkDemo
{
    // Demonstrates how to create a workbook with three worksheets, assign each sheet's A1 cell as a bookmark destination, build a root PdfBookmarkEntry with two child entries, enable ExportDocumentStructure, attach the hierarchy to PdfSaveOptions, and save the workbook as a PDF that shows a clickable outline in any PDF viewer.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add three worksheets
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0]; // default sheet
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");

            // Put sample content in cells that will be bookmark destinations
            sheet1.Cells["A1"].PutValue("Content of Sheet1");
            sheet2.Cells["A1"].PutValue("Content of Sheet2");
            sheet3.Cells["A1"].PutValue("Content of Sheet3");

            // Build the bookmark hierarchy
            PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
            {
                Text = "Workbook Root",
                Destination = sheet1.Cells["A1"],
                IsOpen = true,
                SubEntry = new ArrayList()
            };

            PdfBookmarkEntry subBookmark1 = new PdfBookmarkEntry
            {
                Text = "Sheet2 Section",
                Destination = sheet2.Cells["A1"]
            };

            PdfBookmarkEntry subBookmark2 = new PdfBookmarkEntry
            {
                Text = "Sheet3 Section",
                Destination = sheet3.Cells["A1"]
            };

            // Add sub‑bookmarks to the root entry
            rootBookmark.SubEntry.Add(subBookmark1);
            rootBookmark.SubEntry.Add(subBookmark2);

            // Configure PDF save options:
            // - ExportDocumentStructure = true ensures the outline (bookmarks) is written.
            // - Bookmark = rootBookmark attaches the hierarchy to the PDF.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true,
                Bookmark = rootBookmark
            };

            // Save the workbook as PDF with the defined bookmarks.
            workbook.Save("output_bookmark.pdf", pdfOptions);

            Console.WriteLine("PDF saved with document structure and bookmarks.");
        }
    }
}
