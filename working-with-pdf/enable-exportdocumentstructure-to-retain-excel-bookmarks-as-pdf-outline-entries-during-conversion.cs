using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class ExportBookmarksToPdf
{
    static void Main()
    {
        // Create a new workbook and add worksheets
        Workbook workbook = new Workbook();
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Sheet1";
        sheet1.Cells["A1"].PutValue("Sheet1 Content");

        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
        sheet2.Cells["A1"].PutValue("Sheet2 Content");

        Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");
        sheet3.Cells["A1"].PutValue("Sheet3 Content");

        // Create the root PDF bookmark entry
        PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
        {
            Text = "Root",                     // Title of the root bookmark
            Destination = sheet1.Cells["A1"],  // Destination cell for the root
            IsOpen = true                      // Expand the root entry by default
        };

        // Create sub‑bookmarks for the other worksheets
        PdfBookmarkEntry subBookmark1 = new PdfBookmarkEntry
        {
            Text = "Sheet2",
            Destination = sheet2.Cells["A1"]
        };

        PdfBookmarkEntry subBookmark2 = new PdfBookmarkEntry
        {
            Text = "Sheet3",
            Destination = sheet3.Cells["A1"]
        };

        // Attach sub‑bookmarks to the root entry
        rootBookmark.SubEntry = new ArrayList { subBookmark1, subBookmark2 };

        // Configure PDF save options:
        // - ExportDocumentStructure = true retains the bookmark outline in the PDF
        // - Bookmark = rootBookmark defines the bookmark hierarchy
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ExportDocumentStructure = true,
            Bookmark = rootBookmark
        };

        // Save the workbook as a PDF with the defined bookmarks
        workbook.Save("BookmarksOutput.pdf", pdfOptions);
    }
}