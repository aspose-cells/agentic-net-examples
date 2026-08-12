// Title: Add PDF Bookmarks for Each Worksheet Using Aspose.Cells PdfSaveOptions (C#)
// Description: Demonstrates how to create a hierarchical PdfBookmarkEntry, assign it to PdfSaveOptions.Bookmark, optionally enable ExportDocumentStructure, and save a workbook so every worksheet appears as a clickable PDF bookmark.
// Keywords: Aspose.Cells PDF bookmarks | PdfSaveOptions Bookmark C# | PdfBookmarkEntry example | export Excel to PDF with bookmarks | C# Aspose.Cells PDF outline | Workbook to PDF with navigation | ExportDocumentStructure Aspose
// Common Searches: Aspose.Cells add PDF bookmarks for worksheets | PdfSaveOptions Bookmark property C# example | Create PDF outline from Excel sheets Aspose | Export workbook to PDF with clickable bookmarks .NET | How to generate PDF bookmarks from Excel using Aspose
// Developer Intent: Generate a PDF from an Excel workbook where each worksheet is represented by a PDF bookmark for easy navigation.
// Use Cases: Produce multi‑sheet PDF reports with a navigable outline for readers. | Automate documentation where each Excel tab maps to a PDF chapter. | Enhance large PDF exports from Excel with clickable bookmarks for faster access.
// AI Prompts: Write C# code that iterates through all worksheets in a workbook and builds a PdfBookmarkEntry tree for PdfSaveOptions. | Show how to enable ExportDocumentStructure when saving a workbook to PDF with Aspose.Cells. | Explain the difference between setting PdfSaveOptions.Bookmark and leaving it null.

using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to create a hierarchical PdfBookmarkEntry, assign it to PdfSaveOptions.Bookmark, optionally enable ExportDocumentStructure, and save a workbook so every worksheet appears as a clickable PDF bookmark.
class Program
{
    static void Main()
    {
        // Create a new workbook and add three worksheets
        Workbook workbook = new Workbook();

        // First worksheet (already exists)
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Sheet1";
        sheet1.Cells["A1"].PutValue("Content of Sheet1");

        // Additional worksheets
        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
        sheet2.Cells["A1"].PutValue("Content of Sheet2");

        Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");
        sheet3.Cells["A1"].PutValue("Content of Sheet3");

        // Create a root bookmark entry
        PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
        {
            Text = "Workbook",
            Destination = sheet1.Cells["A1"],
            IsOpen = true
        };

        // Create bookmark entries for each worksheet
        PdfBookmarkEntry bm1 = new PdfBookmarkEntry
        {
            Text = sheet1.Name,
            Destination = sheet1.Cells["A1"]
        };

        PdfBookmarkEntry bm2 = new PdfBookmarkEntry
        {
            Text = sheet2.Name,
            Destination = sheet2.Cells["A1"]
        };

        PdfBookmarkEntry bm3 = new PdfBookmarkEntry
        {
            Text = sheet3.Name,
            Destination = sheet3.Cells["A1"]
        };

        // Attach the worksheet bookmarks as sub‑entries of the root
        rootBookmark.SubEntry = new ArrayList { bm1, bm2, bm3 };

        // Configure PDF save options to include the bookmarks
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Bookmark = rootBookmark,
            ExportDocumentStructure = true   // optional: retain document structure
        };

        // Save the workbook as a PDF with the defined bookmarks
        workbook.Save("WorkbookWithBookmarks.pdf", pdfOptions);
    }
}
