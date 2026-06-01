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
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DataSheet";

            // Fill some data in the range that will be bookmarked
            sheet.Cells["B2"].PutValue("Start of Range");
            sheet.Cells["B3"].PutValue(100);
            sheet.Cells["B4"].PutValue(200);
            sheet.Cells["B5"].PutValue(300);
            sheet.Cells["B6"].PutValue("End of Range");

            // Define a named range that covers B2:B6
            int nameIndex = workbook.Worksheets.Names.Add("MyRange");
            Name namedRange = workbook.Worksheets.Names[nameIndex];
            // The RefersTo string must start with '=' and include the sheet name
            namedRange.RefersTo = $"=DataSheet!$B$2:$B$6";

            // Create a PDF bookmark entry that uses the named destination
            PdfBookmarkEntry bookmark = new PdfBookmarkEntry
            {
                Text = "My Named Range Bookmark",
                // Set DestinationName to the name of the range; this creates a named destination
                DestinationName = "MyRange",
                // Optionally set Destination to the first cell of the range (not required when using DestinationName)
                Destination = sheet.Cells["B2"],
                IsOpen = true
            };

            // If you want sub‑bookmarks, you can add them to the SubEntry collection
            // Example (optional):
            // PdfBookmarkEntry subBookmark = new PdfBookmarkEntry
            // {
            //     Text = "Sub Section",
            //     Destination = sheet.Cells["B4"]
            // };
            // bookmark.SubEntry = new ArrayList { subBookmark };

            // Configure PDF save options with the bookmark
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Bookmark = bookmark
            };

            // Save the workbook as a PDF; the bookmark will navigate to the named range
            workbook.Save("PdfWithNamedRangeBookmark.pdf", pdfOptions);
        }
    }
}