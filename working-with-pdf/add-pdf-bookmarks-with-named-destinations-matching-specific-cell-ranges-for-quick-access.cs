using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class PdfBookmarkDemo
{
    static void Main()
    {
        // Create a new workbook and add worksheets
        Workbook workbook = new Workbook();
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Summary";
        Worksheet sheet2 = workbook.Worksheets.Add("Details");
        Worksheet sheet3 = workbook.Worksheets.Add("Charts");

        // Populate cells that will serve as bookmark destinations
        sheet1.Cells["A1"].PutValue("Summary Page");
        sheet2.Cells["B2"].PutValue("Details Section");
        sheet3.Cells["C3"].PutValue("Charts Overview");
        sheet3.Cells["C5"].PutValue("Chart 1 Data");

        // Create the root bookmark entry
        PdfBookmarkEntry root = new PdfBookmarkEntry
        {
            Text = "Workbook",
            Destination = sheet1.Cells["A1"],
            IsOpen = true,
            SubEntry = new ArrayList()
        };

        // First child bookmark (Summary)
        PdfBookmarkEntry bookmark1 = new PdfBookmarkEntry
        {
            Text = "Summary",
            Destination = sheet1.Cells["A1"],
            DestinationName = "SummaryDest"
        };

        // Second child bookmark (Details)
        PdfBookmarkEntry bookmark2 = new PdfBookmarkEntry
        {
            Text = "Details",
            Destination = sheet2.Cells["B2"],
            DestinationName = "DetailsDest"
        };

        // Third child bookmark (Charts) with a sub‑bookmark
        PdfBookmarkEntry bookmark3 = new PdfBookmarkEntry
        {
            Text = "Charts",
            Destination = sheet3.Cells["C3"],
            DestinationName = "ChartsDest",
            SubEntry = new ArrayList()
        };

        // Sub‑bookmark under Charts
        PdfBookmarkEntry subChart = new PdfBookmarkEntry
        {
            Text = "Chart 1",
            Destination = sheet3.Cells["C5"],
            DestinationName = "Chart1Dest"
        };
        bookmark3.SubEntry.Add(subChart);

        // Assemble the hierarchy
        root.SubEntry.Add(bookmark1);
        root.SubEntry.Add(bookmark2);
        root.SubEntry.Add(bookmark3);

        // Configure PDF save options with the bookmark hierarchy
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Bookmark = root,
            ExportDocumentStructure = true // retain document structure for accessibility
        };

        // Save the workbook as a PDF with bookmarks
        workbook.Save("WorkbookWithBookmarks.pdf", pdfOptions);
    }
}