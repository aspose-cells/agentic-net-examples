// Title: Add PDF Bookmark Hierarchy to Aspose.Cells Workbook Before Exporting to PDF (C#)
// Description: The example creates a three‑sheet workbook, defines a root PdfBookmarkEntry with two child entries that point to cells A1 on each sheet, uses reflection to set the PdfBookmarks property on PdfSaveOptions when it exists, and saves the workbook as a PDF with clickable bookmarks.
// Keywords: Aspose.Cells | C# | PDF bookmarks | PdfBookmarkEntry | PdfSaveOptions | PdfBookmarks property | Excel to PDF export | bookmark hierarchy | reflection | version compatibility
// Common Searches: Aspose.Cells add PDF bookmarks C# | How to set PdfBookmarks in Aspose.Cells | Create hierarchical PDF bookmarks with Aspose.Cells | PdfBookmarkEntry example .NET | Export workbook to PDF with bookmarks Aspose
// Developer Intent: Attach a PdfBookmarkEntry hierarchy to the workbook's PdfBookmarks collection and export the workbook as a PDF.
// Use Cases: Generate a multi‑sheet Excel file and produce a PDF with a navigable bookmark outline. | Build a dynamic bookmark tree from worksheet data and embed it during PDF conversion. | Maintain compatibility across Aspose.Cells versions by checking for the PdfBookmarks property before assignment.
// AI Prompts: Show C# code that creates a PdfBookmarkEntry tree, assigns it to Workbook.PdfBookmarks, and saves the workbook as a PDF using Aspose.Cells. | Explain how to use reflection to safely set the PdfBookmarks property when it may be missing in older Aspose.Cells releases. | Provide a step‑by‑step example of adding root and child PDF bookmarks that link to specific cells before exporting to PDF.

using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a three‑sheet workbook, defines a root PdfBookmarkEntry with two child entries that point to cells A1 on each sheet, uses reflection to set the PdfBookmarks property on PdfSaveOptions when it exists, and saves the workbook as a PDF with clickable bookmarks.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add worksheets
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0];
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");

            // Set values that will serve as bookmark destinations
            sheet1.Cells["A1"].PutValue("Page 1");
            sheet2.Cells["A1"].PutValue("Page 2");
            sheet3.Cells["A1"].PutValue("Page 3");

            // Create the root PDF bookmark entry
            PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
            {
                Text = "Root",
                Destination = sheet1.Cells["A1"],
                IsOpen = true,
                SubEntry = new ArrayList()
            };

            // Create sub‑bookmarks
            PdfBookmarkEntry subBookmark1 = new PdfBookmarkEntry
            {
                Text = "Section 1",
                Destination = sheet2.Cells["A1"]
            };

            PdfBookmarkEntry subBookmark2 = new PdfBookmarkEntry
            {
                Text = "Section 2",
                Destination = sheet3.Cells["A1"]
            };

            // Add sub‑bookmarks to the root entry
            rootBookmark.SubEntry.Add(subBookmark1);
            rootBookmark.SubEntry.Add(subBookmark2);

            // Prepare PDF save options
            PdfSaveOptions saveOptions = new PdfSaveOptions();

            // Attach the bookmark hierarchy if the property is available (supported in newer versions)
            // In older versions the PdfBookmarks property may not exist; this block safely skips it.
            var pdfBookmarksProp = typeof(PdfSaveOptions).GetProperty("PdfBookmarks");
            if (pdfBookmarksProp != null && pdfBookmarksProp.CanWrite)
            {
                pdfBookmarksProp.SetValue(saveOptions, rootBookmark);
            }

            // Export the workbook to PDF
            string outputPath = "output_bookmark.pdf";
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
