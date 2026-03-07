using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class PdfBookmarkDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add worksheets that will contain bookmark destinations
        Worksheet sheet1 = workbook.Worksheets.Add("Intro");
        Worksheet sheet2 = workbook.Worksheets.Add("Details");
        Worksheet sheet3 = workbook.Worksheets.Add("Summary");

        // Populate cells that will serve as bookmark targets
        sheet1.Cells["A1"].PutValue("Introduction");
        sheet2.Cells["A1"].PutValue("Detailed Information");
        sheet3.Cells["A1"].PutValue("Summary");

        // Create the root bookmark entry
        PdfBookmarkEntry root = new PdfBookmarkEntry
        {
            Text = "Document",
            Destination = sheet1.Cells["A1"],          // Link to Intro!A1
            DestinationName = "IntroDest",            // Named destination
            IsOpen = true,                            // Expanded by default
            SubEntry = new ArrayList()                // Container for child entries
        };

        // First sub‑bookmark
        PdfBookmarkEntry sub1 = new PdfBookmarkEntry
        {
            Text = "Details Section",
            Destination = sheet2.Cells["A1"],         // Link to Details!A1
            DestinationName = "DetailsDest"
        };

        // Second sub‑bookmark
        PdfBookmarkEntry sub2 = new PdfBookmarkEntry
        {
            Text = "Summary Section",
            Destination = sheet3.Cells["A1"],         // Link to Summary!A1
            DestinationName = "SummaryDest"
        };

        // Attach sub‑bookmarks to the root entry
        root.SubEntry.Add(sub1);
        root.SubEntry.Add(sub2);

        // Configure PDF save options with the bookmark hierarchy
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Bookmark = root,
            ExportDocumentStructure = true            // Preserve document structure
        };

        // Save the workbook as a PDF file with bookmarks
        workbook.Save("ReportWithBookmarks.pdf", pdfOptions);
    }
}