// Title: Add PDF Bookmarks for Each Worksheet with Aspose.Cells PdfSaveOptions (C#)
// Description: Shows how to build a hidden root PdfBookmarkEntry, attach sub‑bookmarks for every worksheet, enable ExportDocumentStructure, and save the workbook as a PDF where each sheet is a clickable bookmark.
// Keywords: Aspose.Cells PDF bookmarks | PdfSaveOptions Bookmark property | C# add worksheet bookmarks PDF | ExportDocumentStructure Aspose.Cells | Excel to PDF with bookmarks | Aspose.Cells .NET PDF navigation | PdfBookmarkEntry example | Create PDF outline from worksheets
// Common Searches: Aspose.Cells add PDF bookmarks per worksheet | PdfSaveOptions Bookmark C# example | How to generate PDF outline from Excel sheets using Aspose.Cells | Export Excel workbook to PDF with navigation bookmarks | Create hidden root PDF bookmark Aspose.Cells
// Developer Intent: Create a PDF from an Excel workbook where each worksheet appears as an individual bookmark.
// Use Cases: Generate multi‑chapter PDF reports that mirror Excel sheet sections. | Provide end‑users with quick navigation in PDF versions of financial models. | Automate document pipelines that require a PDF outline reflecting workbook tabs. | Build searchable PDFs with a structured bookmark tree for compliance documentation.
// AI Prompts: Write C# code using Aspose.Cells to add a hidden root PdfBookmarkEntry, create sub‑bookmarks for all worksheets, and save the workbook as a PDF. | Explain how ExportDocumentStructure influences PDF bookmark creation with PdfSaveOptions. | Show how to iterate through worksheets programmatically and build a PdfBookmarkEntry hierarchy for PDF export.

using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfBookmarksDemo
{
    // Shows how to build a hidden root PdfBookmarkEntry, attach sub‑bookmarks for every worksheet, enable ExportDocumentStructure, and save the workbook as a PDF where each sheet is a clickable bookmark.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add three worksheets with sample data
            Worksheet sheet1 = workbook.Worksheets[0]; // default first sheet
            sheet1.Name = "Sheet1";
            sheet1.Cells["A1"].PutValue("Content of Sheet1");

            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            sheet2.Cells["A1"].PutValue("Content of Sheet2");

            Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");
            sheet3.Cells["A1"].PutValue("Content of Sheet3");

            // Create the root bookmark (can be hidden by leaving Text empty)
            PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
            {
                Text = "", // empty text hides the root, children appear at top level
                Destination = sheet1.Cells["A1"], // optional, can be any cell
                IsOpen = true
            };

            // Create individual bookmarks for each worksheet
            PdfBookmarkEntry bookmark1 = new PdfBookmarkEntry
            {
                Text = "Sheet1",
                Destination = sheet1.Cells["A1"]
            };

            PdfBookmarkEntry bookmark2 = new PdfBookmarkEntry
            {
                Text = "Sheet2",
                Destination = sheet2.Cells["A1"]
            };

            PdfBookmarkEntry bookmark3 = new PdfBookmarkEntry
            {
                Text = "Sheet3",
                Destination = sheet3.Cells["A1"]
            };

            // Attach the worksheet bookmarks as sub‑entries of the root
            rootBookmark.SubEntry = new ArrayList { bookmark1, bookmark2, bookmark3 };

            // Configure PDF save options with the bookmark hierarchy
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Bookmark = rootBookmark,
                ExportDocumentStructure = true // ensure structure is exported
            };

            // Save the workbook to PDF with bookmarks
            workbook.Save("WorkbookWithBookmarks.pdf", pdfOptions);
        }
    }
}
