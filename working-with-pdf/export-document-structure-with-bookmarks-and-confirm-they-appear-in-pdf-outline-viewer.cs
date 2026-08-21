// Title: Create hierarchical PDF bookmarks from Excel worksheets using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to build a PdfBookmarkEntry tree, enable ExportDocumentStructure, and save an Excel workbook as a PDF with a clickable outline that appears in PDF viewers.
// Keywords: Aspose.Cells PDF bookmarks C# | ExportDocumentStructure Aspose | PdfSaveOptions bookmark hierarchy | Excel to PDF outline Aspose.Cells | PdfBookmarkEntry example
// Common Searches: Aspose.Cells add PDF bookmarks from Excel | export document structure PDF Aspose .NET | C# create nested PDF outline from worksheets | how to show bookmarks in PDF saved by Aspose.Cells | Aspose.Cells PdfSaveOptions ExportDocumentStructure usage
// Developer Intent: Generate a PDF from an Excel workbook that includes a multi‑level bookmark outline for easy navigation in PDF readers.
// Use Cases: Produce a PDF report where each worksheet is a top‑level bookmark. | Create a user manual PDF with sections linked to specific cells. | Export a financial model to PDF with an expandable bookmark tree for auditors.
// AI Prompts: Show how to add sub‑sub bookmarks to the PDF using Aspose.Cells. | Provide code to programmatically verify that the bookmarks appear in the PDF outline. | Explain how to set a bookmark destination to a cell range instead of a single cell.

using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfBookmarkDemo
{
    // Demonstrates how to build a PdfBookmarkEntry tree, enable ExportDocumentStructure, and save an Excel workbook as a PDF with a clickable outline that appears in PDF viewers.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add three worksheets
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0]; // default sheet
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");

            // Populate cells that will serve as bookmark destinations
            sheet1.Cells["A1"].PutValue("Content of Sheet 1");
            sheet2.Cells["A1"].PutValue("Content of Sheet 2");
            sheet3.Cells["A1"].PutValue("Content of Sheet 3");

            // Build bookmark hierarchy
            PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
            {
                Text = "Workbook Root",
                Destination = sheet1.Cells["A1"],
                IsOpen = true,
                SubEntry = new ArrayList()
            };

            PdfBookmarkEntry bookmarkSheet2 = new PdfBookmarkEntry
            {
                Text = "Sheet 2",
                Destination = sheet2.Cells["A1"]
            };

            PdfBookmarkEntry bookmarkSheet3 = new PdfBookmarkEntry
            {
                Text = "Sheet 3",
                Destination = sheet3.Cells["A1"]
            };

            // Add child bookmarks to the root
            rootBookmark.SubEntry.Add(bookmarkSheet2);
            rootBookmark.SubEntry.Add(bookmarkSheet3);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Export the document structure so that the outline (bookmarks) is visible in PDF viewers
                ExportDocumentStructure = true,
                // Attach the bookmark hierarchy
                Bookmark = rootBookmark
            };

            // Save the workbook as PDF with the defined options
            workbook.Save("WorkbookWithBookmarks.pdf", pdfOptions);

            Console.WriteLine("PDF saved with document structure and bookmarks.");
        }
    }
}
