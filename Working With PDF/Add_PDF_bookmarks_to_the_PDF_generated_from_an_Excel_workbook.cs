using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using System.Collections;

class PdfBookmarkExample
{
    static void Main()
    {
        // Load an existing Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access or create worksheets to be used as bookmark destinations
        Worksheet sheet1 = workbook.Worksheets[0];
        Worksheet sheet2 = workbook.Worksheets.Count > 1 ? workbook.Worksheets[1] : workbook.Worksheets.Add("Sheet2");
        Worksheet sheet3 = workbook.Worksheets.Count > 2 ? workbook.Worksheets[2] : workbook.Worksheets.Add("Sheet3");

        // Ensure the destination cells contain some content
        sheet1.Cells["A1"].PutValue("First Sheet Content");
        sheet2.Cells["A1"].PutValue("Second Sheet Content");
        sheet3.Cells["A1"].PutValue("Third Sheet Content");

        // Create the root PDF bookmark entry
        PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
        {
            Text = "Workbook",
            Destination = sheet1.Cells["A1"],
            IsOpen = true
        };

        // Create sub‑bookmarks for the other sheets
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

        // Configure PDF save options with the bookmark hierarchy
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Bookmark = rootBookmark,
            ExportDocumentStructure = true // optional: retain document structure
        };

        // Save the workbook as a PDF file with the defined bookmarks
        workbook.Save("output.pdf", pdfOptions);
    }
}