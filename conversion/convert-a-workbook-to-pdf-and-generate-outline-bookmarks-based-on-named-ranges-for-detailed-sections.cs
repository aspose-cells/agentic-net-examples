using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsBookmarkPdfDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Report";

                // Populate some sample data
                sheet.Cells["A1"].PutValue("Executive Summary");
                sheet.Cells["A2"].PutValue("This is the executive summary.");
                sheet.Cells["A5"].PutValue("Section 1");
                sheet.Cells["A6"].PutValue("Details of section 1...");
                sheet.Cells["A10"].PutValue("Section 2");
                sheet.Cells["A11"].PutValue("Details of section 2...");

                // Create the root bookmark entry
                PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
                {
                    Text = "Report",
                    Destination = sheet.Cells["A1"], // Destination for the root (optional)
                    IsOpen = true,
                    SubEntry = new ArrayList()
                };

                // Helper to create a bookmark entry for a specific cell
                PdfBookmarkEntry CreateBookmark(string text, string cellRef)
                {
                    return new PdfBookmarkEntry
                    {
                        Text = text,
                        Destination = sheet.Cells[cellRef]
                    };
                }

                // Add sub‑bookmarks for each section
                rootBookmark.SubEntry.Add(CreateBookmark("Executive Summary", "A1"));
                rootBookmark.SubEntry.Add(CreateBookmark("Section 1", "A5"));
                rootBookmark.SubEntry.Add(CreateBookmark("Section 2", "A10"));

                // Configure PDF save options with the bookmark hierarchy
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    Bookmark = rootBookmark,
                    ExportDocumentStructure = true // Preserve document structure for accessibility
                };

                // Save the workbook as a PDF
                string outputPath = "ReportWithBookmarks.pdf";
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"PDF saved successfully with outline bookmarks to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}