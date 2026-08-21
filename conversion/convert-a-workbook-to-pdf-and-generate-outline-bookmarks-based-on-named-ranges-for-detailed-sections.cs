// Title: Convert Excel to PDF with hierarchical outline bookmarks from named ranges using Aspose.Cells for .NET
// Description: This C# example demonstrates how to export an Aspose.Cells workbook to PDF while building a multi‑level bookmark outline. Named ranges (or worksheet titles) are mapped to PdfBookmarkEntry objects, the bookmark tree is attached to PdfSaveOptions, and the PDF is saved with document structure enabled. The code also ensures the output directory exists and removes any previous file before writing the new PDF.
// Keywords: Aspose.Cells PDF bookmarks | C# Excel to PDF conversion | named ranges PDF outline | PdfBookmarkEntry example | PdfSaveOptions bookmark | export Excel workbook as PDF | outline navigation in PDF | Aspose.Cells .NET | programmatic PDF generation | document structure PDF
// Common Searches: how to add PDF bookmarks when converting Excel to PDF with Aspose.Cells | C# create hierarchical PDF outline from named ranges in Excel | Aspose.Cells save workbook as PDF with custom bookmark tree | export Excel worksheets to PDF with navigation bookmarks .NET | PdfBookmarkEntry usage example Aspose.Cells
// Developer Intent: Generate a PDF from an Excel workbook and automatically include a hierarchical bookmark outline that links to each named range or worksheet.
// Use Cases: Produce navigable PDF reports where each section (e.g., Introduction, Chapter, Conclusion) is reachable via outline bookmarks. | Automate documentation pipelines that convert Excel templates into PDF manuals with a clickable table of contents. | Integrate PDF generation with bookmarks into CI/CD workflows for consistent, searchable deliverables.
// AI Prompts: Show C# code that creates a PDF bookmark hierarchy from named ranges in an Aspose.Cells workbook. | Explain how to configure PdfSaveOptions to embed a custom bookmark tree and preserve document structure when saving to PDF. | Provide a snippet that checks for the output folder, deletes an existing PDF, and saves the new file with bookmarks using Aspose.Cells.

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// This C# example demonstrates how to export an Aspose.Cells workbook to PDF while building a multi‑level bookmark outline. Named ranges (or worksheet titles) are mapped to PdfBookmarkEntry objects, the bookmark tree is attached to PdfSaveOptions, and the PDF is saved with document structure enabled. The code also ensures the output directory exists and removes any previous file before writing the new PDF.
class WorkbookToPdfWithBookmarks
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // -----------------------------------------------------------------
            // Populate worksheets with sample data
            // -----------------------------------------------------------------
            Worksheet wsIntro = wb.Worksheets[0];
            wsIntro.Name = "Introduction";
            wsIntro.Cells["A1"].PutValue("Introduction");
            wsIntro.Cells["A2"].PutValue("This is the intro section.");

            Worksheet wsChapter = wb.Worksheets.Add("Chapter1");
            wsChapter.Cells["A1"].PutValue("Chapter 1");
            wsChapter.Cells["A2"].PutValue("Details of chapter 1.");

            Worksheet wsConclusion = wb.Worksheets.Add("Conclusion");
            wsConclusion.Cells["A1"].PutValue("Conclusion");
            wsConclusion.Cells["A2"].PutValue("Final remarks.");

            // -----------------------------------------------------------------
            // Build PDF bookmark hierarchy
            // -----------------------------------------------------------------
            PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
            {
                Text = "Document",
                Destination = wsIntro.Cells["A1"], // Root points to first section
                IsOpen = true,
                SubEntry = new ArrayList()
            };

            // Helper to create a bookmark entry from a worksheet
            PdfBookmarkEntry CreateBookmark(string title, Worksheet ws)
            {
                // Destination is the first cell of the worksheet (assumed A1 here)
                Cell dest = ws.Cells["A1"];
                return new PdfBookmarkEntry
                {
                    Text = title,
                    Destination = dest
                };
            }

            // Add sub‑bookmarks for each section
            rootBookmark.SubEntry.Add(CreateBookmark("Introduction", wsIntro));
            rootBookmark.SubEntry.Add(CreateBookmark("Chapter 1", wsChapter));
            rootBookmark.SubEntry.Add(CreateBookmark("Conclusion", wsConclusion));

            // -----------------------------------------------------------------
            // Configure PDF save options with the bookmark tree
            // -----------------------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Bookmark = rootBookmark,
                ExportDocumentStructure = true
            };

            // -----------------------------------------------------------------
            // Save the workbook as PDF
            // -----------------------------------------------------------------
            string outputPath = "DocumentWithBookmarks.pdf";

            try
            {
                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Delete existing file if present
                if (File.Exists(outputPath))
                {
                    File.Delete(outputPath);
                }

                wb.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving PDF: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
