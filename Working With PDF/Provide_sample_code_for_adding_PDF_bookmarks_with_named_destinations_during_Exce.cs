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

            // Set values in cells that will be bookmark destinations
            sheet1.Cells["A1"].PutValue("Content of Sheet1");
            sheet2.Cells["A1"].PutValue("Content of Sheet2");
            sheet3.Cells["A1"].PutValue("Content of Sheet3");

            // Create the root bookmark entry
            PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
            {
                Text = "Root Bookmark",
                Destination = sheet1.Cells["A1"],
                DestinationName = "Sheet1Dest",
                IsOpen = true
            };

            // Create sub‑bookmarks with their own named destinations
            PdfBookmarkEntry subBookmark1 = new PdfBookmarkEntry
            {
                Text = "Go to Sheet2",
                Destination = sheet2.Cells["A1"],
                DestinationName = "Sheet2Dest"
            };

            PdfBookmarkEntry subBookmark2 = new PdfBookmarkEntry
            {
                Text = "Go to Sheet3",
                Destination = sheet3.Cells["A1"],
                DestinationName = "Sheet3Dest"
            };

            // Attach sub‑bookmarks to the root entry
            rootBookmark.SubEntry = new ArrayList { subBookmark1, subBookmark2 };

            // Configure PDF save options with the bookmark hierarchy
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Bookmark = rootBookmark,
                ExportDocumentStructure = true
            };

            // Save the workbook as PDF with bookmarks and named destinations
            workbook.Save("ExcelWithBookmarks.pdf", pdfOptions);
        }
    }
}