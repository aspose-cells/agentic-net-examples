using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class WorkbookToPdfWithBookmarks
{
    static void Main()
    {
        // Create a new workbook and add sample worksheets
        Workbook workbook = new Workbook();
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Sheet1";
        sheet1.Cells["A1"].Value = "Content of Sheet1";

        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
        sheet2.Cells["A1"].Value = "Content of Sheet2";

        Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");
        sheet3.Cells["A1"].Value = "Content of Sheet3";

        // Create the root bookmark (optional, can be hidden by leaving Text null)
        PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
        {
            Text = "Workbook",
            Destination = sheet1.Cells["A1"], // Destination of root can be first sheet
            IsOpen = true,
            SubEntry = new ArrayList()
        };

        // Create a bookmark entry for each worksheet
        foreach (Worksheet ws in workbook.Worksheets)
        {
            PdfBookmarkEntry entry = new PdfBookmarkEntry
            {
                Text = ws.Name,
                Destination = ws.Cells["A1"]
            };
            rootBookmark.SubEntry.Add(entry);
        }

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Export document structure enables outline/bookmarks in PDF
            ExportDocumentStructure = true,
            // Assign the constructed bookmark hierarchy
            Bookmark = rootBookmark
        };

        // Save the workbook as PDF using the options (lifecycle rule)
        workbook.Save("WorkbookWithBookmarks.pdf", pdfOptions);
    }
}