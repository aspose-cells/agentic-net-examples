using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook (default first worksheet is already added)
        Workbook workbook = new Workbook();

        // Rename the default sheet and add two more worksheets
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Sheet1";
        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
        Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");

        // Set a cell in each sheet that will serve as the bookmark destination
        sheet1.Cells["A1"].PutValue("Sheet1 Content");
        sheet2.Cells["A1"].PutValue("Sheet2 Content");
        sheet3.Cells["A1"].PutValue("Sheet3 Content");

        // Create the root bookmark entry
        PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
        {
            Text = "Workbook",               // Title shown in the PDF bookmark pane
            Destination = sheet1.Cells["A1"], // Destination for the root (can be any sheet)
            IsOpen = true                    // Expand the root node by default
        };

        // Create a bookmark entry for each worksheet
        PdfBookmarkEntry bmSheet1 = new PdfBookmarkEntry
        {
            Text = "Sheet1",
            Destination = sheet1.Cells["A1"]
        };
        PdfBookmarkEntry bmSheet2 = new PdfBookmarkEntry
        {
            Text = "Sheet2",
            Destination = sheet2.Cells["A1"]
        };
        PdfBookmarkEntry bmSheet3 = new PdfBookmarkEntry
        {
            Text = "Sheet3",
            Destination = sheet3.Cells["A1"]
        };

        // Attach the sheet bookmarks as sub‑entries of the root
        rootBookmark.SubEntry = new ArrayList { bmSheet1, bmSheet2, bmSheet3 };

        // Configure PDF save options to include the bookmarks
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Bookmark = rootBookmark,
            ExportDocumentStructure = true // optional: retain document structure
        };

        // Save the workbook as a PDF with the defined bookmarks
        workbook.Save("WorkbookWithBookmarks.pdf", pdfOptions);
    }
}