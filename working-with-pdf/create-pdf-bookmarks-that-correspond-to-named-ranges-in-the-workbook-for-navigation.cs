// Title: Add PDF Bookmarks for Named Ranges with Aspose.Cells in C#
// Description: Shows how to build a hidden root PdfBookmarkEntry, generate a bookmark for every simple named range that points to its first cell, attach the hierarchy to PdfSaveOptions, and save the workbook as a PDF with clickable navigation entries.
// Keywords: Aspose.Cells PDF bookmarks | C# PDF bookmark example | named range PDF navigation | PdfSaveOptions bookmark | Excel to PDF outline | AsposeRange bookmark | export workbook to PDF with bookmarks | PDF navigation from Excel
// Common Searches: Aspose.Cells add PDF bookmarks from named ranges | C# create PDF outline for Excel workbook | How to link PDF bookmarks to cells using Aspose.Cells | Generate PDF with navigation bookmarks in C# | Aspose.Cells PdfSaveOptions bookmark hierarchy
// Developer Intent: Create a PDF from an Excel workbook and embed outline bookmarks that jump to each defined named range.
// Use Cases: Enable readers to jump directly to key data sections in a PDF report. | Organize multiple range bookmarks under a hidden root to keep the PDF outline tidy. | Export multi‑sheet workbooks with clickable bookmarks for every named range, improving usability.
// AI Prompts: Modify the example to include page numbers in each PDF bookmark. | Show how to exclude named ranges that span more than one worksheet when building bookmarks. | Rewrite the code to write the PDF to a MemoryStream instead of a file.

using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsPdfBookmarksDemo
{
    // Shows how to build a hidden root PdfBookmarkEntry, generate a bookmark for every simple named range that points to its first cell, attach the hierarchy to PdfSaveOptions, and save the workbook as a PDF with clickable navigation entries.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Sheet1";
                sheet1.Cells["A1"].PutValue("First Sheet");
                sheet1.Cells["B2"].PutValue("Data 1");

                Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
                sheet2.Cells["A1"].PutValue("Second Sheet");
                sheet2.Cells["C3"].PutValue("Data 2");

                // Define named ranges
                sheet1.Cells.CreateRange("A1:B2").Name = "FirstRange";
                sheet2.Cells.CreateRange("C3:D4").Name = "SecondRange";

                // Create a hidden root bookmark (empty Text) to hold all named range entries
                PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
                {
                    Text = "",                     // Hidden root
                    IsOpen = true,
                    SubEntry = new ArrayList()
                };

                // Iterate through all defined names in the workbook
                foreach (Name definedName in workbook.Worksheets.Names)
                {
                    // Retrieve the range the name refers to
                    AsposeRange range = definedName.GetRange();
                    if (range == null) continue; // Skip if not a simple range

                    // Destination cell: first cell of the range
                    Cell destinationCell = range.Worksheet.Cells[range.FirstRow, range.FirstColumn];

                    // Create a bookmark entry for this named range
                    PdfBookmarkEntry entry = new PdfBookmarkEntry
                    {
                        Text = definedName.Text,          // Bookmark title = name of the range
                        Destination = destinationCell,    // Link target
                        IsOpen = true
                    };

                    // Add the entry to the root's sub‑entries
                    rootBookmark.SubEntry.Add(entry);
                }

                // Configure PDF save options with the constructed bookmark hierarchy
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    Bookmark = rootBookmark
                };

                // Save the workbook as a PDF with bookmarks
                workbook.Save("WorkbookWithBookmarks.pdf", pdfOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
