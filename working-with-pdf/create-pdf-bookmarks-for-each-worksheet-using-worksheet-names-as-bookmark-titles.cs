using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class PdfBookmarksExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add worksheets and sample data
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "FirstSheet";
        sheet1.Cells["A1"].PutValue("Content of First Sheet");

        Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");
        sheet2.Cells["A1"].PutValue("Content of Second Sheet");

        Worksheet sheet3 = workbook.Worksheets.Add("ThirdSheet");
        sheet3.Cells["A1"].PutValue("Content of Third Sheet");

        // Create a hidden root bookmark entry; its children will appear as top‑level bookmarks
        PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
        {
            Text = null,               // hide root entry
            SubEntry = new ArrayList(),
            IsOpen = true
        };

        // Create a bookmark for each worksheet using the sheet name as the title
        foreach (Worksheet ws in workbook.Worksheets)
        {
            PdfBookmarkEntry entry = new PdfBookmarkEntry
            {
                Text = ws.Name,               // bookmark title
                Destination = ws.Cells["A1"], // link target
                IsOpen = true
            };
            rootBookmark.SubEntry.Add(entry);
        }

        // Configure PDF save options with the bookmark hierarchy
        PdfSaveOptions options = new PdfSaveOptions
        {
            Bookmark = rootBookmark,
            ExportDocumentStructure = true // ensure bookmarks are written to the PDF
        };

        // Save the workbook as a PDF with the created bookmarks
        workbook.Save("WorksheetsBookmarks.pdf", options);
    }
}