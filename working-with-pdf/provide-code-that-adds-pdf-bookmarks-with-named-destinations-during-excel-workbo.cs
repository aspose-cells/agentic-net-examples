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
            // Create a new workbook (contains a default worksheet)
            Workbook workbook = new Workbook();

            // Use the default worksheet as Sheet1
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";

            // Add two more worksheets
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
                DestinationName = "Sheet1Start",
                IsOpen = true
            };

            // Create sub‑bookmarks with named destinations
            PdfBookmarkEntry subBookmark1 = new PdfBookmarkEntry
            {
                Text = "Go to Sheet2",
                Destination = sheet2.Cells["A1"],
                DestinationName = "Sheet2Start"
            };

            PdfBookmarkEntry subBookmark2 = new PdfBookmarkEntry
            {
                Text = "Go to Sheet3",
                Destination = sheet3.Cells["A1"],
                DestinationName = "Sheet3Start"
            };

            // Attach sub‑bookmarks to the root entry
            rootBookmark.SubEntry = new ArrayList { subBookmark1, subBookmark2 };

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Bookmark = rootBookmark,
                ExportDocumentStructure = true
            };

            // Save the workbook as PDF with the defined bookmarks
            workbook.Save("WorkbookWithBookmarks.pdf", pdfOptions);
        }
    }
}