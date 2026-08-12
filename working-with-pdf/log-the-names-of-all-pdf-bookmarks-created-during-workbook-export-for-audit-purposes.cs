// Title: Audit PDF Bookmark Names When Exporting an Aspose.Cells Workbook to PDF (C#)
// Description: This example creates a workbook with three worksheets, builds a hierarchical PdfBookmarkEntry (root with sub‑bookmarks), recursively logs each bookmark's Text to the console, assigns the hierarchy to PdfSaveOptions.Bookmark, and saves the workbook as a PDF while handling errors. Use it to capture an audit trail of all PDF bookmarks generated during export.
// Keywords: Aspose.Cells PDF bookmarks | C# PdfBookmarkEntry recursion | log PDF bookmark titles | audit PDF bookmarks Aspose | export workbook to PDF with bookmarks | PdfSaveOptions Bookmark | Aspose.Cells example C# | PDF bookmark hierarchy
// Common Searches: how to list PDF bookmarks created by Aspose.Cells | C# code to log bookmark titles during PDF export | traverse PdfBookmarkEntry hierarchy Aspose.Cells | audit PDF bookmarks after workbook.Save | Aspose.Cells export workbook to PDF with bookmarks
// Developer Intent: Capture and record the titles of every PDF bookmark generated while saving a workbook to PDF with Aspose.Cells.
// Use Cases: Create an audit log of bookmark names before distributing the PDF for compliance verification. | Generate documentation that enumerates all worksheet bookmarks included in the exported PDF. | Perform quality‑assurance checks to ensure expected worksheets appear as PDF bookmarks.
// AI Prompts: Provide a method that returns all PDF bookmark titles as a List<string> instead of printing them. | Show how to write bookmark names to a timestamped log file while preserving exception handling. | Explain how to add page numbers to each PdfBookmarkEntry and include that data in the audit log.

using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// This example creates a workbook with three worksheets, builds a hierarchical PdfBookmarkEntry (root with sub‑bookmarks), recursively logs each bookmark's Text to the console, assigns the hierarchy to PdfSaveOptions.Bookmark, and saves the workbook as a PDF while handling errors. Use it to capture an audit trail of all PDF bookmarks generated during export.
class PdfBookmarkLogger
{
    // Recursively logs bookmark titles (Text) to the console.
    static void LogBookmarks(PdfBookmarkEntry entry)
    {
        if (entry == null) return;

        if (!string.IsNullOrEmpty(entry.Text))
        {
            Console.WriteLine($"Bookmark: {entry.Text}");
        }

        if (entry.SubEntry != null)
        {
            foreach (PdfBookmarkEntry child in entry.SubEntry)
            {
                LogBookmarks(child);
            }
        }
    }

    static void Main()
    {
        try
        {
            // Create a new workbook and clear the default worksheet.
            Workbook workbook = new Workbook();
            workbook.Worksheets.Clear();

            // Add three worksheets with unique names.
            Worksheet sheet1 = workbook.Worksheets.Add("Sheet1");
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");

            // Set values in cells that will be bookmark destinations.
            sheet1.Cells["A1"].PutValue("Sheet1 Content");
            sheet2.Cells["A1"].PutValue("Sheet2 Content");
            sheet3.Cells["A1"].PutValue("Sheet3 Content");

            // Create root bookmark.
            PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
            {
                Text = "Root",
                Destination = sheet1.Cells["A1"],
                IsOpen = true
            };

            // Create sub‑bookmarks.
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

            // Attach sub‑bookmarks to the root.
            rootBookmark.SubEntry = new ArrayList { subBookmark1, subBookmark2 };

            // Log all bookmark names for audit.
            LogBookmarks(rootBookmark);

            // Configure PDF save options with the bookmark hierarchy.
            PdfSaveOptions options = new PdfSaveOptions
            {
                Bookmark = rootBookmark
            };

            // Save the workbook as PDF.
            workbook.Save("output.pdf", options);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
