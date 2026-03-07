using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfBookmarkDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and clear the default worksheet
            Workbook workbook = new Workbook();
            workbook.Worksheets.Clear();

            // Add three worksheets
            Worksheet sheet1 = workbook.Worksheets.Add("Sheet1");
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");

            // Set values in cells that will serve as bookmark destinations
            sheet1.Cells["A1"].PutValue("Content of Sheet1");
            sheet2.Cells["A1"].PutValue("Content of Sheet2");
            sheet3.Cells["A1"].PutValue("Content of Sheet3");

            // Create the root PDF bookmark entry with a named destination
            PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
            {
                Text = "Root",
                Destination = sheet1.Cells["A1"],
                DestinationName = "RootDest",
                IsOpen = true
            };

            // Create sub‑bookmarks, each with its own named destination
            PdfBookmarkEntry subBookmark1 = new PdfBookmarkEntry
            {
                Text = "Sheet2 Section",
                Destination = sheet2.Cells["A1"],
                DestinationName = "Sheet2Dest"
            };

            PdfBookmarkEntry subBookmark2 = new PdfBookmarkEntry
            {
                Text = "Sheet3 Section",
                Destination = sheet3.Cells["A1"],
                DestinationName = "Sheet3Dest"
            };

            // Attach sub‑entries to the root bookmark
            rootBookmark.SubEntry = new ArrayList { subBookmark1, subBookmark2 };

            // Configure PDF save options with the bookmark hierarchy
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Bookmark = rootBookmark,
                ExportDocumentStructure = true
            };

            // Save the workbook as a PDF file with bookmarks and named destinations
            workbook.Save("WorkbookWithBookmarks.pdf", pdfOptions);
        }
    }
}