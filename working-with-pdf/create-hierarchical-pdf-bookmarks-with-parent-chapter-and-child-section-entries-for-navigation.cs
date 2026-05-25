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
            // Create a new workbook and add three worksheets (chapters)
            Workbook workbook = new Workbook();
            Worksheet chapter1 = workbook.Worksheets[0];
            chapter1.Name = "Chapter 1";
            Worksheet chapter2 = workbook.Worksheets.Add("Chapter 2");
            Worksheet chapter3 = workbook.Worksheets.Add("Chapter 3");

            // Populate a cell in each worksheet that will serve as the bookmark destination
            chapter1.Cells["A1"].PutValue("Content of Chapter 1");
            chapter2.Cells["A1"].PutValue("Content of Chapter 2");
            chapter3.Cells["A1"].PutValue("Content of Chapter 3");

            // ----- Build hierarchical PDF bookmarks -----
            // Root bookmark (invisible, its children will appear at top level)
            PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
            {
                Text = null,                     // Invisible root
                SubEntry = new ArrayList(),
                IsOpen = true
            };

            // Chapter 1 bookmark (parent)
            PdfBookmarkEntry chapter1Bookmark = new PdfBookmarkEntry
            {
                Text = "Chapter 1",
                Destination = chapter1.Cells["A1"],
                IsOpen = true,
                SubEntry = new ArrayList()
            };

            // Sections under Chapter 1
            PdfBookmarkEntry section1_1 = new PdfBookmarkEntry
            {
                Text = "Section 1.1",
                Destination = chapter1.Cells["B1"] // Example destination; can be any cell
            };
            PdfBookmarkEntry section1_2 = new PdfBookmarkEntry
            {
                Text = "Section 1.2",
                Destination = chapter1.Cells["C1"]
            };
            chapter1Bookmark.SubEntry.Add(section1_1);
            chapter1Bookmark.SubEntry.Add(section1_2);

            // Chapter 2 bookmark (parent) with no sub‑sections
            PdfBookmarkEntry chapter2Bookmark = new PdfBookmarkEntry
            {
                Text = "Chapter 2",
                Destination = chapter2.Cells["A1"],
                IsOpen = true
            };

            // Chapter 3 bookmark (parent) with nested sections
            PdfBookmarkEntry chapter3Bookmark = new PdfBookmarkEntry
            {
                Text = "Chapter 3",
                Destination = chapter3.Cells["A1"],
                IsOpen = true,
                SubEntry = new ArrayList()
            };
            PdfBookmarkEntry section3_1 = new PdfBookmarkEntry
            {
                Text = "Section 3.1",
                Destination = chapter3.Cells["B1"]
            };
            PdfBookmarkEntry section3_2 = new PdfBookmarkEntry
            {
                Text = "Section 3.2",
                Destination = chapter3.Cells["C1"]
            };
            chapter3Bookmark.SubEntry.Add(section3_1);
            chapter3Bookmark.SubEntry.Add(section3_2);

            // Assemble the hierarchy under the invisible root
            rootBookmark.SubEntry.Add(chapter1Bookmark);
            rootBookmark.SubEntry.Add(chapter2Bookmark);
            rootBookmark.SubEntry.Add(chapter3Bookmark);

            // Configure PDF save options with the bookmark hierarchy
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Bookmark = rootBookmark,
                ExportDocumentStructure = true   // Preserve structure for better navigation
            };

            // Save the workbook as PDF with hierarchical bookmarks
            workbook.Save("HierarchicalBookmarks.pdf", pdfOptions);

            Console.WriteLine("PDF with hierarchical bookmarks created successfully.");
        }
    }
}