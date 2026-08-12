// Title: C# – Create PDF Bookmarks from Named Ranges with Aspose.Cells
// Description: Demonstrates how to build a PDF outline from Excel named ranges. The code creates a workbook, defines three named ranges, generates a PdfBookmarkEntry hierarchy that points each bookmark to the range's first cell, attaches the hierarchy to PdfSaveOptions, and saves the file as a navigable PDF.
// Keywords: Aspose.Cells PDF bookmarks | C# named range PDF navigation | export Excel to PDF with outline | Aspose.Cells PdfBookmarkEntry example | PDF table of contents from Excel
// Common Searches: Aspose.Cells add PDF bookmarks from named ranges | C# export workbook to PDF with bookmarks | How to create PDF outline using Aspose.Cells | Generate clickable PDF table of contents from Excel | Aspose.Cells PDF bookmark hierarchy tutorial
// Developer Intent: Generate a PDF where each defined name in the workbook appears as a clickable bookmark.
// Use Cases: Provide a quick‑jump table of contents in financial PDFs (SalesData, ExpensesData, SummaryData). | Enable end‑users to navigate large multi‑sheet reports without scrolling. | Automate PDF report generation with a structured bookmark tree for regulatory submissions.
// AI Prompts: Write C# code that iterates over all defined names in an Aspose.Cells workbook and creates PdfBookmarkEntry objects for each. | Explain how to group named‑range bookmarks under sheet‑level parent entries when exporting to PDF with Aspose.Cells. | Suggest best‑practice error handling for saving a workbook as PDF with a bookmark hierarchy using Aspose.Cells.

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsPdfBookmarksDemo
{
    // Demonstrates how to build a PDF outline from Excel named ranges. The code creates a workbook, defines three named ranges, generates a PdfBookmarkEntry hierarchy that points each bookmark to the range's first cell, attaches the hierarchy to PdfSaveOptions, and saves the file as a navigable PDF.
    class Program
    {
        static void Main()
        {
            try
            {
                // -------------------- Create workbook and sample data --------------------
                Workbook workbook = new Workbook();

                // Add three worksheets
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Sales";
                sheet1.Cells["A1"].PutValue("Quarter 1");
                sheet1.Cells["A2"].PutValue(1200);
                sheet1.Cells["B1"].PutValue("Quarter 2");
                sheet1.Cells["B2"].PutValue(1500);

                Worksheet sheet2 = workbook.Worksheets.Add("Expenses");
                sheet2.Cells["A1"].PutValue("Q1");
                sheet2.Cells["A2"].PutValue(300);
                sheet2.Cells["B1"].PutValue("Q2");
                sheet2.Cells["B2"].PutValue(400);

                Worksheet sheet3 = workbook.Worksheets.Add("Summary");
                sheet3.Cells["A1"].PutValue("Total Sales");
                sheet3.Cells["A2"].PutValue(2700);
                sheet3.Cells["B1"].PutValue("Total Expenses");
                sheet3.Cells["B2"].PutValue(700);

                // -------------------- Define named ranges --------------------
                AsposeRange salesRange = sheet1.Cells.CreateRange("A1:B2");
                salesRange.Name = "SalesData";

                AsposeRange expensesRange = sheet2.Cells.CreateRange("A1:B2");
                expensesRange.Name = "ExpensesData";

                AsposeRange summaryRange = sheet3.Cells.CreateRange("A1:B2");
                summaryRange.Name = "SummaryData";

                // -------------------- Build PDF bookmark hierarchy --------------------
                PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
                {
                    Text = "Workbook Bookmarks",
                    IsOpen = true,
                    SubEntry = new ArrayList()
                };

                // Iterate through all defined names in the workbook
                foreach (Name definedName in workbook.Worksheets.Names)
                {
                    // Obtain the range the name refers to
                    AsposeRange rng = definedName.GetRange();
                    if (rng == null) continue; // Skip if the name does not refer to a range

                    // Determine the worksheet that contains the range
                    Worksheet targetSheet;
                    if (definedName.SheetIndex == -1) // Global name
                    {
                        targetSheet = rng.Worksheet; // Worksheet that owns the range
                    }
                    else // Sheet‑specific name
                    {
                        targetSheet = workbook.Worksheets[definedName.SheetIndex];
                    }

                    // Destination cell – top‑left cell of the range
                    Cell destinationCell = targetSheet.Cells[rng.FirstRow, rng.FirstColumn];

                    // Create a bookmark entry for this named range
                    PdfBookmarkEntry entry = new PdfBookmarkEntry
                    {
                        Text = definedName.Text,          // Bookmark title = name of the range
                        Destination = destinationCell,    // Link points to the first cell of the range
                        IsOpen = true                     // Expand this entry by default
                    };

                    // Add the entry to the root's sub‑entries
                    rootBookmark.SubEntry.Add(entry);
                }

                // -------------------- Configure PDF save options with the bookmark --------------------
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    Bookmark = rootBookmark
                };

                // -------------------- Save workbook as PDF --------------------
                string outputPath = "WorkbookWithBookmarks.pdf";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));

                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                try
                {
                    workbook.Save(outputPath, pdfOptions);
                    Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine("Failed to save PDF: " + saveEx.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
