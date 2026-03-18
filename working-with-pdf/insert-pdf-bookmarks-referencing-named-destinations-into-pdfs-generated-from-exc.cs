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
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add three worksheets
            Worksheet sheet1 = workbook.Worksheets[0]; // default sheet
            sheet1.Name = "Sheet1";
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");

            // Put sample content that will serve as bookmark destinations
            sheet1.Cells["A1"].PutValue("Content of Sheet1");
            sheet2.Cells["A1"].PutValue("Content of Sheet2");
            sheet3.Cells["A1"].PutValue("Content of Sheet3");

            // Create the root bookmark entry
            PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
            {
                Text = "Root Bookmark",
                Destination = sheet1.Cells["A1"],
                DestinationName = "Sheet1Dest", // named destination
                IsOpen = true,
                SubEntry = new ArrayList()
            };

            // Create sub‑bookmark for Sheet2
            PdfBookmarkEntry subBookmark1 = new PdfBookmarkEntry
            {
                Text = "Sheet2 Section",
                Destination = sheet2.Cells["A1"],
                DestinationName = "Sheet2Dest"
            };

            // Create sub‑bookmark for Sheet3
            PdfBookmarkEntry subBookmark2 = new PdfBookmarkEntry
            {
                Text = "Sheet3 Section",
                Destination = sheet3.Cells["A1"],
                DestinationName = "Sheet3Dest"
            };

            // Add sub‑bookmarks to the root entry
            rootBookmark.SubEntry.Add(subBookmark1);
            rootBookmark.SubEntry.Add(subBookmark2);

            // Configure PDF save options with the bookmark hierarchy
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Bookmark = rootBookmark,
                ExportDocumentStructure = true // ensures the bookmark structure is written
            };

            // Save the workbook as a PDF with bookmarks
            workbook.Save("WorkbookWithBookmarks.pdf", pdfOptions);
        }
    }
}