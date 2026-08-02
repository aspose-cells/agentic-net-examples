// Title: Create PDF Bookmarks with Named Destinations for Excel Cell Ranges using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to generate a PDF from an Excel workbook, define a hidden root bookmark, add child PdfBookmarkEntry objects linked to specific cell ranges (A1, A10, A20, A30), assign named destinations, enable a hierarchical bookmark tree, preserve document structure for accessibility, and save the file as a navigable PDF.
// Keywords: Aspose.Cells | PDF bookmarks | named destinations | C# | PdfBookmarkEntry | PdfSaveOptions | ExportDocumentStructure | Excel to PDF | bookmark hierarchy | .NET
// Common Searches: Aspose.Cells add PDF bookmarks from Excel cells | C# create named destinations in PDF using Aspose.Cells | How to build hierarchical PDF bookmarks with Aspose.Cells .NET | Export Excel worksheet to PDF with accessible bookmarks | Set PdfSaveOptions bookmark hierarchy Aspose.Cells
// Developer Intent: Generate a PDF from an Excel workbook that contains a structured bookmark outline, where each bookmark jumps to a defined cell range via a named destination.
// Use Cases: Produce a multi‑section report PDF where users can click bookmarks to jump to Executive Summary, Financial Overview, Operational Details, and Conclusion. | Enhance navigation in large exported spreadsheets for accessibility tools by linking PDF bookmarks to named destinations. | Create a custom bookmark tree in the PDF output to organize content for end‑users without modifying the original Excel file.
// AI Prompts: Show C# code that adds PDF bookmarks with named destinations for specific Excel cell ranges using Aspose.Cells. | Generate a hidden root PdfBookmarkEntry and child bookmarks linked to cells A1, A10, A20, A30, then save as PDF with ExportDocumentStructure enabled. | Explain how to customize the PDF bookmark hierarchy and assign DestinationName values in Aspose.Cells for .NET.

using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfBookmarks
{
    // Demonstrates how to generate a PDF from an Excel workbook, define a hidden root bookmark, add child PdfBookmarkEntry objects linked to specific cell ranges (A1, A10, A20, A30), assign named destinations, enable a hierarchical bookmark tree, preserve document structure for accessibility, and save the file as a navigable PDF.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add a worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Report";

            // Populate some sample data in different ranges
            sheet.Cells["A1"].PutValue("Executive Summary");
            sheet.Cells["A10"].PutValue("Financial Overview");
            sheet.Cells["A20"].PutValue("Operational Details");
            sheet.Cells["A30"].PutValue("Conclusion");

            // ----- Create PDF bookmark hierarchy -----
            // Root bookmark (hidden, so its children appear at top level)
            PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
            {
                Text = null,                     // Hidden root
                SubEntry = new ArrayList(),
                IsOpen = true
            };

            // Bookmark for Executive Summary (range A1:A5)
            PdfBookmarkEntry execSummary = new PdfBookmarkEntry
            {
                Text = "Executive Summary",
                Destination = sheet.Cells["A1"],
                DestinationName = "ExecSummary", // Named destination
                IsOpen = true
            };

            // Bookmark for Financial Overview (range A10:A15)
            PdfBookmarkEntry financial = new PdfBookmarkEntry
            {
                Text = "Financial Overview",
                Destination = sheet.Cells["A10"],
                DestinationName = "FinancialOverview",
                IsOpen = true
            };

            // Bookmark for Operational Details (range A20:A25)
            PdfBookmarkEntry operational = new PdfBookmarkEntry
            {
                Text = "Operational Details",
                Destination = sheet.Cells["A20"],
                DestinationName = "OperationalDetails",
                IsOpen = true
            };

            // Bookmark for Conclusion (range A30:A35)
            PdfBookmarkEntry conclusion = new PdfBookmarkEntry
            {
                Text = "Conclusion",
                Destination = sheet.Cells["A30"],
                DestinationName = "Conclusion",
                IsOpen = true
            };

            // Assemble the hierarchy
            rootBookmark.SubEntry.Add(execSummary);
            rootBookmark.SubEntry.Add(financial);
            rootBookmark.SubEntry.Add(operational);
            rootBookmark.SubEntry.Add(conclusion);

            // Configure PDF save options with the bookmark structure
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Bookmark = rootBookmark,
                ExportDocumentStructure = true   // Preserve document structure for accessibility
            };

            // Save the workbook as PDF with the defined bookmarks
            workbook.Save("ReportWithBookmarks.pdf", pdfOptions);
        }
    }
}
