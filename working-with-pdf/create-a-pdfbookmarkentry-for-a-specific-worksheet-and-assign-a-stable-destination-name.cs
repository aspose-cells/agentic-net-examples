// Title: Create PDF Bookmarks with Stable Destination Names in Aspose.Cells (C#)
// Description: Demonstrates how to generate a PDF from an Aspose.Cells workbook, assign a meaningful worksheet name, add content, create a PdfBookmarkEntry with a stable DestinationName, add a sub‑bookmark, configure PdfSaveOptions with the bookmark hierarchy, and save the file as a PDF containing expandable bookmarks.
// Keywords: Aspose.Cells | PdfBookmarkEntry | PDF bookmarks | DestinationName | named destination | C# | .NET | PdfSaveOptions | sub‑bookmark | hierarchical bookmarks | worksheet cell bookmark | Aspose.Cells PDF export
// Common Searches: Aspose.Cells set named destination for PDF bookmark | C# create hierarchical PDF bookmarks with Aspose.Cells | PdfBookmarkEntry DestinationName example | How to add sub‑bookmarks to PDF using Aspose.Cells | Save workbook as PDF with bookmarks Aspose.Cells .NET
// Developer Intent: Add stable, named PDF bookmarks to a workbook and export it as a PDF using Aspose.Cells for .NET.
// Use Cases: Generate a report PDF where the bookmarks jump directly to overview and analysis sections on the same worksheet. | Create expandable PDF bookmarks with consistent destination names for automated report generation. | Add nested bookmarks to large PDFs to improve navigation and user experience.
// AI Prompts: Show C# code that creates a PdfBookmarkEntry with a DestinationName pointing to a worksheet cell and saves the workbook as a PDF using Aspose.Cells. | Provide an example of adding sub‑bookmarks to a PDF bookmark hierarchy with stable named destinations in Aspose.Cells for .NET. | Explain how to configure PdfSaveOptions to include an expandable bookmark structure when exporting a workbook to PDF.

using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfBookmarkDemo
{
    // Demonstrates how to generate a PDF from an Aspose.Cells workbook, assign a meaningful worksheet name, add content, create a PdfBookmarkEntry with a stable DestinationName, add a sub‑bookmark, configure PdfSaveOptions with the bookmark hierarchy, and save the file as a PDF containing expandable bookmarks.
    public class CreateBookmarkWithDestinationName
    {
        // Entry point required for console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and give it a meaningful name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Report";

            // Add sample content that will serve as the bookmark destination
            sheet.Cells["A1"].PutValue("Report Overview");
            sheet.Cells["A5"].PutValue("Detailed Analysis");

            // Create a PDF bookmark entry pointing to cell A1
            PdfBookmarkEntry bookmark = new PdfBookmarkEntry
            {
                Text = "Overview Section",                 // Title shown in PDF bookmarks pane
                Destination = sheet.Cells["A1"],           // Cell that the bookmark links to
                DestinationName = "ReportOverview",       // Stable named destination
                IsOpen = true                              // Expand this bookmark by default
            };

            // Optionally add a sub‑bookmark pointing to another cell on the same sheet
            PdfBookmarkEntry subBookmark = new PdfBookmarkEntry
            {
                Text = "Analysis Section",
                Destination = sheet.Cells["A5"],
                DestinationName = "ReportAnalysis"
            };
            // SubEntry expects a collection of PdfBookmarkEntry objects
            bookmark.SubEntry = new ArrayList { subBookmark };

            // Configure PDF save options with the bookmark hierarchy
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Bookmark = bookmark
            };

            // Save the workbook as a PDF file with the defined bookmarks
            workbook.Save("ReportWithBookmarks.pdf", pdfOptions);
        }
    }
}
