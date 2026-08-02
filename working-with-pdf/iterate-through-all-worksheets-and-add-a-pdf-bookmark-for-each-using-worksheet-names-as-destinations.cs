using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfBookmarks
{
    // Author: Aspose.Cells .NET example – adds a PDF bookmark for each worksheet using its name.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample worksheets (replace with loading if needed)
            Workbook workbook = new Workbook();
            workbook.Worksheets.Add("Finance");
            workbook.Worksheets.Add("HR");
            workbook.Worksheets.Add("Engineering");

            // Populate a cell in each sheet to serve as the bookmark destination
            foreach (Worksheet ws in workbook.Worksheets)
            {
                ws.Cells["A1"].PutValue($"{ws.Name} Content");
            }

            // Root bookmark – empty text so its children appear at the top level
            PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
            {
                Text = null,
                SubEntry = new ArrayList()
            };

            // Create a bookmark entry for each worksheet
            foreach (Worksheet ws in workbook.Worksheets)
            {
                PdfBookmarkEntry sheetBookmark = new PdfBookmarkEntry
                {
                    Text = ws.Name,               // Bookmark title
                    Destination = ws.Cells["A1"] // Destination cell
                };

                // Add to root's sub‑entries
                rootBookmark.SubEntry.Add(sheetBookmark);
            }

            // Configure PDF save options with the constructed bookmark hierarchy
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Bookmark = rootBookmark
            };

            // Save the workbook as PDF with bookmarks
            workbook.Save("WorksheetsWithBookmarks.pdf", pdfOptions);
        }
    }
}