using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace PdfBookmarkDemo
{
    // Author: Aspose.Cells .NET example
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add three worksheets
            Workbook workbook = new Workbook();
            Worksheet chapterSheet = workbook.Worksheets[0];
            Worksheet sectionSheet1 = workbook.Worksheets.Add("Section1");
            Worksheet sectionSheet2 = workbook.Worksheets.Add("Section2");

            // Set values that will serve as bookmark destinations
            chapterSheet.Cells["A1"].PutValue("Chapter 1 Content");
            sectionSheet1.Cells["A1"].PutValue("Section 1.1 Content");
            sectionSheet2.Cells["A1"].PutValue("Section 1.2 Content");

            // ----- Build hierarchical PDF bookmarks -----
            // Root bookmark representing the chapter
            PdfBookmarkEntry chapterBookmark = new PdfBookmarkEntry
            {
                Text = "Chapter 1",
                Destination = chapterSheet.Cells["A1"],
                IsOpen = true,          // Expanded by default
                SubEntry = new ArrayList()
            };

            // Child bookmark for Section 1.1
            PdfBookmarkEntry sectionBookmark1 = new PdfBookmarkEntry
            {
                Text = "Section 1.1",
                Destination = sectionSheet1.Cells["A1"]
            };

            // Child bookmark for Section 1.2
            PdfBookmarkEntry sectionBookmark2 = new PdfBookmarkEntry
            {
                Text = "Section 1.2",
                Destination = sectionSheet2.Cells["A1"]
            };

            // Attach child entries to the chapter bookmark
            chapterBookmark.SubEntry.Add(sectionBookmark1);
            chapterBookmark.SubEntry.Add(sectionBookmark2);

            // Configure PDF save options with the bookmark hierarchy
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Bookmark = chapterBookmark
            };

            // Save the workbook as a PDF with hierarchical bookmarks
            workbook.Save("HierarchicalBookmarks.pdf", pdfOptions);
        }
    }
}