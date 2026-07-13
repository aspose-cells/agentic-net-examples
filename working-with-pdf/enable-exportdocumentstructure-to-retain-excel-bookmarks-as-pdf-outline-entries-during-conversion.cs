using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and add three worksheets
        Workbook workbook = new Workbook();
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Sheet1";
        sheet1.Cells["A1"].PutValue("Content of Sheet1");

        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
        sheet2.Cells["A1"].PutValue("Content of Sheet2");

        Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");
        sheet3.Cells["A1"].PutValue("Content of Sheet3");

        // Build the PDF bookmark hierarchy
        PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
        {
            Text = "Workbook Root",
            Destination = sheet1.Cells["A1"],
            IsOpen = true,
            SubEntry = new ArrayList()
        };

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

        // Attach sub‑bookmarks to the root
        rootBookmark.SubEntry.Add(subBookmark1);
        rootBookmark.SubEntry.Add(subBookmark2);

        // Configure PDF save options: enable document structure export and assign bookmarks
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Bookmark = rootBookmark,
            ExportDocumentStructure = true // retain Excel bookmarks as PDF outline entries
        };

        // Save the workbook as a PDF with the configured options
        workbook.Save("WorkbookWithBookmarks.pdf", saveOptions);
    }
}

// Author: Aspose.Cells .NET example – retains Excel bookmarks in PDF outline.