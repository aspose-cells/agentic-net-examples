using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfBookmarks
{
    // Author: Aspose.Cells .NET example – creates PDF bookmarks for each worksheet using worksheet names.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample worksheets
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets.Add("Sales");
            Worksheet sheet2 = workbook.Worksheets.Add("Inventory");
            Worksheet sheet3 = workbook.Worksheets.Add("Summary");

            // Put a marker value in each sheet – this cell will be the bookmark destination
            sheet1.Cells["A1"].PutValue("Sales Data");
            sheet2.Cells["A1"].PutValue("Inventory Data");
            sheet3.Cells["A1"].PutValue("Summary Data");

            // Create a root bookmark entry with no text so its children appear at the top level
            PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
            {
                Text = null,
                SubEntry = new ArrayList()
            };

            // Add a bookmark for each worksheet
            AddWorksheetBookmark(rootBookmark, sheet1);
            AddWorksheetBookmark(rootBookmark, sheet2);
            AddWorksheetBookmark(rootBookmark, sheet3);

            // Configure PDF save options with the constructed bookmark hierarchy
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Bookmark = rootBookmark
            };

            // Save the workbook as a PDF with bookmarks
            workbook.Save("WorksheetsWithBookmarks.pdf", pdfOptions);
        }

        // Helper method to create and attach a bookmark for a worksheet
        private static void AddWorksheetBookmark(PdfBookmarkEntry root, Worksheet ws)
        {
            PdfBookmarkEntry entry = new PdfBookmarkEntry
            {
                Text = ws.Name,               // Use worksheet name as bookmark title
                Destination = ws.Cells["A1"] // Link to cell A1 in the sheet
            };
            root.SubEntry.Add(entry);
        }
    }
}