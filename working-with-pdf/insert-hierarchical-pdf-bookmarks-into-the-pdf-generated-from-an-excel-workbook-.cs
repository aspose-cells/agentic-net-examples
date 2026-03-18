using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfBookmarksDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add three worksheets
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets.Add("Introduction");
            Worksheet sheet2 = workbook.Worksheets.Add("Chapter 1");
            Worksheet sheet3 = workbook.Worksheets.Add("Chapter 2");

            // Put some content in each worksheet – these cells will be bookmark destinations
            sheet1.Cells["A1"].PutValue("Welcome to the Introduction");
            sheet2.Cells["A1"].PutValue("Content of Chapter 1");
            sheet3.Cells["A1"].PutValue("Content of Chapter 2");

            // ----- Build hierarchical PDF bookmarks -----
            // Root bookmark (will appear as the top entry in the PDF bookmark pane)
            PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
            {
                Text = "Document Outline",
                Destination = sheet1.Cells["A1"],
                IsOpen = true,               // Expanded by default
                SubEntry = new ArrayList()   // Container for child entries
            };

            // First level child – Chapter 1
            PdfBookmarkEntry chapter1 = new PdfBookmarkEntry
            {
                Text = "Chapter 1",
                Destination = sheet2.Cells["A1"]
            };

            // Second level child under Chapter 1 – Section 1.1
            PdfBookmarkEntry section11 = new PdfBookmarkEntry
            {
                Text = "Section 1.1",
                Destination = sheet2.Cells["A2"] // Assuming A2 holds section start
            };

            // Second level child under Chapter 1 – Section 1.2
            PdfBookmarkEntry section12 = new PdfBookmarkEntry
            {
                Text = "Section 1.2",
                Destination = sheet2.Cells["A3"] // Assuming A3 holds section start
            };

            // Add sections to Chapter 1
            chapter1.SubEntry = new ArrayList { section11, section12 };

            // First level child – Chapter 2
            PdfBookmarkEntry chapter2 = new PdfBookmarkEntry
            {
                Text = "Chapter 2",
                Destination = sheet3.Cells["A1"]
            };

            // Add Chapter 1 and Chapter 2 to the root bookmark
            rootBookmark.SubEntry.Add(chapter1);
            rootBookmark.SubEntry.Add(chapter2);

            // Configure PDF save options with the bookmark hierarchy
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Bookmark = rootBookmark,
                ExportDocumentStructure = true   // Preserve the document structure for accessibility
            };

            // Save the workbook as a PDF with hierarchical bookmarks
            workbook.Save("HierarchicalBookmarks.pdf", pdfOptions);

            Console.WriteLine("PDF with hierarchical bookmarks created successfully.");
        }
    }
}