// Title: Add a PDF bookmark to the first visible cell with a named destination using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a Workbook, set values in cells A1 and B2, define a PdfBookmarkEntry that targets cell A1, assign a named destination ("FirstCell"), enable the bookmark to open on load, configure PdfSaveOptions to retain document structure, and save the workbook as a PDF containing the bookmark.
// Keywords: Aspose.Cells PDF bookmark C# | named destination PDF Aspose.Cells | PdfBookmarkEntry example | ExportDocumentStructure PDF | bookmark first visible cell | Aspose.Cells save as PDF with bookmark
// Common Searches: Aspose.Cells add PDF bookmark to a cell | C# create PDF bookmark with named destination | Export Excel to PDF with bookmarks using Aspose | PdfSaveOptions bookmark example .NET | How to set PDF bookmark to first visible cell
// Developer Intent: Generate a PDF from an Excel workbook that includes a bookmark pointing to the first visible cell, using a named destination for external linking.
// Use Cases: Quick navigation to a summary cell in generated PDF reports. | Providing stable anchors for cross‑document references in PDFs. | Creating interactive PDFs with predefined navigation points for end users.
// AI Prompts: Show how to add multiple PDF bookmarks for different cells with Aspose.Cells in C#. | Explain the role of ExportDocumentStructure in preserving PDF bookmarks and how to verify them in a viewer. | Provide code to retrieve the named destination of a PDF bookmark created with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfBookmarkDemo
{
    // Demonstrates how to create a Workbook, set values in cells A1 and B2, define a PdfBookmarkEntry that targets cell A1, assign a named destination ("FirstCell"), enable the bookmark to open on load, configure PdfSaveOptions to retain document structure, and save the workbook as a PDF containing the bookmark.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data (optional, just to have visible content)
            sheet.Cells["A1"].PutValue("First visible cell");
            sheet.Cells["B2"].PutValue("Another cell");

            // Create a PDF bookmark entry
            PdfBookmarkEntry bookmark = new PdfBookmarkEntry
            {
                // Title shown in PDF bookmarks pane
                Text = "First Visible Cell",
                // Destination cell (first visible cell)
                Destination = sheet.Cells["A1"],
                // Optional named destination (useful for external references)
                DestinationName = "FirstCell",
                // Expand the bookmark when PDF is opened
                IsOpen = true
            };

            // Configure PDF save options with the bookmark
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Bookmark = bookmark,
                // Export document structure so that bookmarks are retained
                ExportDocumentStructure = true
            };

            // Save the workbook as PDF with the configured bookmark
            workbook.Save("FirstCellBookmark.pdf", pdfOptions);
        }
    }
}
