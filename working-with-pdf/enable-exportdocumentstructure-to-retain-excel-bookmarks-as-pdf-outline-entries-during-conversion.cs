// Title: Create PDF Outline from Excel Bookmarks using Aspose.Cells ExportDocumentStructure (C#)
// Description: A C# demo that creates a workbook with three worksheets, defines a root PdfBookmarkEntry and sub‑entries linked to cell A1 of each sheet, enables PdfSaveOptions.ExportDocumentStructure, assigns the bookmark hierarchy, and saves the file as a PDF whose outline mirrors the Excel bookmarks.
// Keywords: Aspose.Cells | ExportDocumentStructure | PDF bookmarks | C# | PdfSaveOptions | PdfBookmarkEntry | Excel to PDF outline | bookmark hierarchy | save workbook as PDF | Aspose.Cells GitHub example | document structure PDF
// Common Searches: Aspose.Cells export Excel bookmarks to PDF outline | how to enable ExportDocumentStructure in PdfSaveOptions C# | create PDF bookmark hierarchy from workbook cells | Aspose.Cells PDF outline entries example | preserve sheet navigation when converting Excel to PDF
// Developer Intent: Generate a PDF from an Excel workbook that includes an expandable bookmark outline linked to specific cells.
// Use Cases: Add a top‑level PDF bookmark named "Workbook Outline" with child entries for each worksheet. | Link each bookmark to a designated cell (e.g., A1) so clicking the outline jumps to the corresponding sheet. | Control the initial open/closed state of the outline by setting the IsOpen property on PdfBookmarkEntry objects. | Produce PDFs with navigable outlines for reports, manuals, or multi‑sheet dashboards.
// AI Prompts: Show C# code that builds a nested PDF bookmark structure from an Aspose.Cells workbook and saves it with ExportDocumentStructure enabled. | Explain how PdfSaveOptions.Bookmark works together with ExportDocumentStructure to create PDF outlines. | Give an example of linking PDF bookmarks to cell A1 on different worksheets using Aspose.Cells.

using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsBookmarkPdfDemo
{
    // A C# demo that creates a workbook with three worksheets, defines a root PdfBookmarkEntry and sub‑entries linked to cell A1 of each sheet, enables PdfSaveOptions.ExportDocumentStructure, assigns the bookmark hierarchy, and saves the file as a PDF whose outline mirrors the Excel bookmarks.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add three worksheets
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");

            // Populate cells that will serve as bookmark destinations
            sheet1.Cells["A1"].PutValue("Content of Sheet1");
            sheet2.Cells["A1"].PutValue("Content of Sheet2");
            sheet3.Cells["A1"].PutValue("Content of Sheet3");

            // Create root bookmark entry
            PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
            {
                Text = "Workbook Outline",
                Destination = sheet1.Cells["A1"],
                IsOpen = true,
                SubEntry = new ArrayList()
            };

            // Create sub‑bookmarks for the other sheets
            PdfBookmarkEntry subBookmark1 = new PdfBookmarkEntry
            {
                Text = "Sheet2",
                Destination = sheet2.Cells["A1"]
            };

            PdfBookmarkEntry subBookmark2 = new PdfBookmarkEntry
            {
                Text = "Sheet3",
                Destination = sheet3.Cells["A1"]
            };

            // Add sub‑bookmarks to the root entry
            rootBookmark.SubEntry.Add(subBookmark1);
            rootBookmark.SubEntry.Add(subBookmark2);

            // Configure PDF save options:
            // - ExportDocumentStructure = true enables PDF outline (bookmarks)
            // - Bookmark = rootBookmark defines the outline hierarchy
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true,
                Bookmark = rootBookmark
            };

            // Save the workbook as PDF with the specified options
            workbook.Save("WorkbookWithBookmarks.pdf", pdfSaveOptions);

            Console.WriteLine("PDF saved with document structure and bookmarks.");
        }
    }
}
