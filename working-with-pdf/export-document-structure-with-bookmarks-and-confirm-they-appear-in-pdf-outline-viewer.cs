// Title: Export an Excel workbook to PDF with a custom bookmark outline using Aspose.Cells for .NET (C#)
// Common Searches: how to generate PDF bookmarks from Excel worksheets using Aspose.Cells C# | Aspose.Cells PdfSaveOptions ExportDocumentStructure example for creating PDF outline | C# code to add hierarchical bookmarks to a PDF saved from a workbook | save each Excel sheet as a separate page in PDF with Aspose.Cells and include bookmarks | verify PDF outline (bookmarks) after exporting Excel with Aspose.Cells .NET
// Tags: Aspose.Cells export workbook to PDF with bookmarks | PdfSaveOptions ExportDocumentStructure C# | PdfBookmarkEntry hierarchy Aspose.Cells | one page per worksheet PDF Aspose.Cells | Excel to PDF outline bookmarks .NET

using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace ExportDocumentStructureWithBookmarks
{
    // The sample creates a three‑sheet workbook, fills cells, builds a PdfBookmarkEntry hierarchy (root, Introduction, Financials with Revenue and Expense sub‑bookmarks, and Conclusion), configures PdfSaveOptions to export the document structure and assign the bookmark tree, enables one page per sheet, saves the file as DocumentWithBookmarks.pdf, and writes console output confirming the PDF creation and bookmark presence.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a workbook and populate it with sample data
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();

                // Add three worksheets
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Introduction";
                Worksheet sheet2 = workbook.Worksheets.Add("Financials");
                Worksheet sheet3 = workbook.Worksheets.Add("Conclusion");

                // Fill some cells that will be used as bookmark destinations
                sheet1.Cells["A1"].PutValue("Introduction Section");
                sheet2.Cells["A1"].PutValue("Financial Overview");
                sheet3.Cells["A1"].PutValue("Final Thoughts");

                // Add a few more rows just to make the PDF pages a bit larger
                for (int i = 2; i <= 30; i++)
                {
                    sheet1.Cells[$"A{i}"].PutValue($"Intro line {i - 1}");
                    sheet2.Cells[$"A{i}"].PutValue($"Finance line {i - 1}");
                    sheet3.Cells[$"A{i}"].PutValue($"Conclusion line {i - 1}");
                }

                // -----------------------------------------------------------------
                // 2. Build the PDF bookmark hierarchy
                // -----------------------------------------------------------------
                // Root bookmark (will appear as the top entry in the PDF outline)
                PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
                {
                    Text = "Document Outline",
                    Destination = sheet1.Cells["A1"], // point to the first sheet
                    IsOpen = true,
                    SubEntry = new ArrayList()
                };

                // First child – Introduction
                PdfBookmarkEntry introBookmark = new PdfBookmarkEntry
                {
                    Text = "Introduction",
                    Destination = sheet1.Cells["A1"],
                    IsOpen = true
                };

                // Second child – Financials (collapsed by default)
                PdfBookmarkEntry financialsBookmark = new PdfBookmarkEntry
                {
                    Text = "Financials",
                    Destination = sheet2.Cells["A1"],
                    IsOpen = false,
                    IsCollapse = true,
                    SubEntry = new ArrayList()
                };

                // Sub‑bookmarks under Financials
                PdfBookmarkEntry revenueBookmark = new PdfBookmarkEntry
                {
                    Text = "Revenue Details",
                    Destination = sheet2.Cells["A5"]
                };
                PdfBookmarkEntry expenseBookmark = new PdfBookmarkEntry
                {
                    Text = "Expense Details",
                    Destination = sheet2.Cells["A15"]
                };
                financialsBookmark.SubEntry.Add(revenueBookmark);
                financialsBookmark.SubEntry.Add(expenseBookmark);

                // Third child – Conclusion
                PdfBookmarkEntry conclusionBookmark = new PdfBookmarkEntry
                {
                    Text = "Conclusion",
                    Destination = sheet3.Cells["A1"],
                    IsOpen = true
                };

                // Assemble the hierarchy
                rootBookmark.SubEntry.Add(introBookmark);
                rootBookmark.SubEntry.Add(financialsBookmark);
                rootBookmark.SubEntry.Add(conclusionBookmark);

                // -----------------------------------------------------------------
                // 3. Configure PDF save options
                // -----------------------------------------------------------------
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Export the document structure (bookmarks) into the PDF outline
                    ExportDocumentStructure = true,

                    // Attach the bookmark hierarchy we built above
                    Bookmark = rootBookmark,

                    // Optional: one page per worksheet for clearer navigation
                    OnePagePerSheet = true
                };

                // -----------------------------------------------------------------
                // 4. Save the workbook as PDF
                // -----------------------------------------------------------------
                string outputPath = "DocumentWithBookmarks.pdf";
                workbook.Save(outputPath, pdfOptions);

                // -----------------------------------------------------------------
                // 5. Confirmation
                // -----------------------------------------------------------------
                if (System.IO.File.Exists(outputPath))
                {
                    Console.WriteLine($"PDF successfully created at '{outputPath}'.");
                    Console.WriteLine("The PDF contains a document outline (bookmarks) as defined in the code.");
                    Console.WriteLine("Open the PDF in any viewer that supports outlines (e.g., Adobe Acrobat Reader) to verify.");
                }
                else
                {
                    Console.WriteLine("Failed to create the PDF file.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while exporting the document structure with bookmarks:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
