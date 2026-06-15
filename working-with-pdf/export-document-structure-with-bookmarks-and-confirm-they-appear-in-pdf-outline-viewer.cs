using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsBookmarkExport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add three worksheets
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0]; // default sheet
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");

            // Populate cells that will serve as bookmark destinations
            sheet1.Cells["A1"].PutValue("Content of Sheet 1");
            sheet2.Cells["A1"].PutValue("Content of Sheet 2");
            sheet3.Cells["A1"].PutValue("Content of Sheet 3");

            // Build bookmark hierarchy
            PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
            {
                Text = "Workbook Root",
                Destination = sheet1.Cells["A1"],
                IsOpen = true,
                SubEntry = new ArrayList()
            };

            PdfBookmarkEntry subBookmark1 = new PdfBookmarkEntry
            {
                Text = "Sheet2 Section",
                Destination = sheet2.Cells["A1"]
            };

            PdfBookmarkEntry subBookmark2 = new PdfBookmarkEntry
            {
                Text = "Sheet3 Section",
                Destination = sheet3.Cells["A1"]
            };

            // Add sub‑bookmarks to the root entry
            rootBookmark.SubEntry.Add(subBookmark1);
            rootBookmark.SubEntry.Add(subBookmark2);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Export the document structure so that bookmarks are included in the PDF outline
                ExportDocumentStructure = true,
                // Assign the bookmark hierarchy
                Bookmark = rootBookmark
            };

            // Save the workbook as PDF
            string outputPath = "WorkbookWithBookmarks.pdf";
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF saved to '{outputPath}'.");
            Console.WriteLine("Bookmarks have been added. Open the PDF in a viewer that shows the outline (e.g., Adobe Reader) to verify.");
        }
    }
}