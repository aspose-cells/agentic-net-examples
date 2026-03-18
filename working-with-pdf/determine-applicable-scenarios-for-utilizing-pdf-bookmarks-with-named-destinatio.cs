using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class PdfBookmarkScenarios
{
    static void Main()
    {
        // ------------------------------------------------------------
        // Create a workbook with three worksheets to demonstrate bookmarks
        // ------------------------------------------------------------
        Workbook wb = new Workbook();

        // Worksheet 0 – will act as the summary page
        Worksheet ws1 = wb.Worksheets[0];
        ws1.Name = "Summary";
        ws1.Cells["A1"].PutValue("Summary Page");

        // Worksheet 1 – details page
        Worksheet ws2 = wb.Worksheets.Add("Details");
        ws2.Cells["A1"].PutValue("Details Page");

        // Worksheet 2 – appendix page
        Worksheet ws3 = wb.Worksheets.Add("Appendix");
        ws3.Cells["A1"].PutValue("Appendix Page");

        // ------------------------------------------------------------
        // Scenario 1: Simple hierarchical bookmarks using cell destinations
        // ------------------------------------------------------------
        PdfBookmarkEntry root = new PdfBookmarkEntry
        {
            Text = "Document Overview",          // Visible root entry
            Destination = ws1.Cells["A1"],       // Link to Summary!A1
            IsOpen = true                        // Expand this node in PDF viewer
        };

        PdfBookmarkEntry details = new PdfBookmarkEntry
        {
            Text = "Details Section",
            Destination = ws2.Cells["A1"]        // Link to Details!A1
        };

        PdfBookmarkEntry appendix = new PdfBookmarkEntry
        {
            Text = "Appendix Section",
            Destination = ws3.Cells["A1"]        // Link to Appendix!A1
        };

        // Attach sub‑entries to the root
        root.SubEntry = new ArrayList { details, appendix };

        // ------------------------------------------------------------
        // Scenario 2: Using named destinations (DestinationName)
        // These names can be referenced from external PDFs or from
        // within the same PDF using link annotations.
        // ------------------------------------------------------------
        PdfBookmarkEntry namedRoot = new PdfBookmarkEntry
        {
            Text = "Named Root",
            Destination = ws1.Cells["A1"],
            DestinationName = "SummaryDest",     // Define a named destination
            IsOpen = true
        };

        PdfBookmarkEntry namedDetails = new PdfBookmarkEntry
        {
            Text = "Named Details",
            Destination = ws2.Cells["A1"],
            DestinationName = "DetailsDest"
        };

        // ------------------------------------------------------------
        // Scenario 3: Hidden root entry (Text = null) – children are
        // inserted at the current level, flattening the hierarchy.
        // ------------------------------------------------------------
        PdfBookmarkEntry hiddenRoot = new PdfBookmarkEntry
        {
            Text = null,                         // Hidden node
            SubEntry = new ArrayList { namedRoot, namedDetails }
        };

        // ------------------------------------------------------------
        // Combine all scenarios into a single master bookmark tree
        // ------------------------------------------------------------
        PdfBookmarkEntry masterRoot = new PdfBookmarkEntry
        {
            Text = "Master Bookmark",
            Destination = ws1.Cells["A1"],
            IsOpen = true,
            SubEntry = new ArrayList { root, hiddenRoot }
        };

        // ------------------------------------------------------------
        // Configure PDF save options
        // ExportDocumentStructure must be true for the bookmark tree
        // to be written into the PDF file.
        // ------------------------------------------------------------
        PdfSaveOptions options = new PdfSaveOptions
        {
            ExportDocumentStructure = true,
            Bookmark = masterRoot
        };

        // ------------------------------------------------------------
        // Save the workbook as PDF with the defined bookmarks
        // ------------------------------------------------------------
        wb.Save("ExcelWithBookmarks.pdf", options);
    }
}