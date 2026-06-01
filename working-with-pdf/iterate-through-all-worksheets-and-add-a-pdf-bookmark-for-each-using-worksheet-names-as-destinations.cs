using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfBookmarks
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Add sample worksheets for demonstration
            workbook.Worksheets.Add("Finance");
            workbook.Worksheets.Add("HR");
            workbook.Worksheets.Add("IT");

            // Ensure each worksheet has a cell to serve as a bookmark destination
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.Cells["A1"].PutValue($"{sheet.Name} Content");
            }

            // Create a root bookmark entry with no text.
            // When Text is null, its children are placed at the top level.
            PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
            {
                Text = null,
                SubEntry = new ArrayList()
            };

            // Iterate through all worksheets and create a bookmark for each.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                PdfBookmarkEntry entry = new PdfBookmarkEntry
                {
                    Text = sheet.Name,                 // Use worksheet name as bookmark title
                    Destination = sheet.Cells["A1"],   // Destination cell for the bookmark
                    IsOpen = true                      // Expand the bookmark by default
                };

                // Add the entry to the root's sub‑entries collection
                rootBookmark.SubEntry.Add(entry);
            }

            // Configure PDF save options with the constructed bookmark hierarchy
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Bookmark = rootBookmark
            };

            // Save the workbook as a PDF with the bookmarks
            workbook.Save("WorkbookWithBookmarks.pdf", pdfOptions);
        }
    }
}