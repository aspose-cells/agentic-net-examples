// Title: Generate PDF with Worksheet‑Based Table of Contents using Aspose.Cells for .NET (C#)
// Description: C# example that creates a workbook, renders each worksheet to a single PDF page, builds a root PdfBookmarkEntry containing sheet names and start pages, and saves the file with ExportDocumentStructure so the bookmarks serve as a clickable table of contents.
// Keywords: Aspose.Cells | C# | PDF bookmarks | table of contents | OnePagePerSheet | PdfSaveOptions | Excel to PDF | document structure | Aspose.Cells for .NET | PDF generation
// Common Searches: Aspose.Cells add PDF bookmarks from Excel sheets | C# create PDF table of contents from workbook | Export Excel worksheets to PDF with page numbers Aspose | How to generate clickable TOC in PDF using Aspose.Cells | One page per sheet PDF Aspose.Cells .NET
// Developer Intent: Create a PDF where each Excel worksheet is rendered on its own page and a clickable TOC lists the sheet names with corresponding page numbers.
// Use Cases: Produce a multi‑sheet PDF report with a navigable TOC for quick access to sections such as Summary, Details, and Statistics. | Generate printable documentation where each Excel sheet becomes a separate PDF page and the TOC reflects exact page numbers for reference. | Automate PDF creation with worksheet‑based bookmarks for integration into document management systems that require searchable PDF structure.
// AI Prompts: Show how to add hierarchical sub‑bookmarks for sections within each worksheet while keeping the main TOC intact. | Provide code to stream the PDF to an ASP.NET response, preserving the generated table of contents. | Explain how to customize the root bookmark title and set individual entries to be collapsed or expanded by default.

using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfTocDemo
{
    // C# example that creates a workbook, renders each worksheet to a single PDF page, builds a root PdfBookmarkEntry containing sheet names and start pages, and saves the file with ExportDocumentStructure so the bookmarks serve as a clickable table of contents.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample worksheets
            Workbook workbook = new Workbook();

            // Worksheet 1
            Worksheet ws1 = workbook.Worksheets[0];
            ws1.Name = "Summary";
            ws1.Cells["A1"].PutValue("Summary Data");
            for (int i = 2; i <= 30; i++)
                ws1.Cells[$"A{i}"].PutValue($"Item {i - 1}");

            // Worksheet 2
            Worksheet ws2 = workbook.Worksheets.Add("Details");
            ws2.Cells["A1"].PutValue("Details Header");
            for (int i = 2; i <= 50; i++)
                ws2.Cells[$"A{i}"].PutValue($"Detail {i - 1}");

            // Worksheet 3
            Worksheet ws3 = workbook.Worksheets.Add("Statistics");
            ws3.Cells["A1"].PutValue("Statistics Header");
            for (int i = 2; i <= 40; i++)
                ws3.Cells[$"A{i}"].PutValue($"Stat {i - 1}");

            // Options for rendering – each sheet will be rendered on a single page
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true
            };

            // Calculate the starting page number for each worksheet
            int currentPage = 1;
            PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
            {
                Text = "Table of Contents",
                IsOpen = true,
                SubEntry = new ArrayList()
            };

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Render the sheet to obtain its page count
                SheetRender sheetRender = new SheetRender(sheet, renderOptions);
                int sheetPages = sheetRender.PageCount; // With OnePagePerSheet this will be 1

                // Create a bookmark entry for the sheet
                PdfBookmarkEntry entry = new PdfBookmarkEntry
                {
                    Text = $"{sheet.Name} - Page {currentPage}",
                    Destination = sheet.Cells["A1"]
                };

                // Add the entry to the root bookmark
                rootBookmark.SubEntry.Add(entry);

                // Update the running page number
                currentPage += sheetPages;

                sheetRender.Dispose();
            }

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true, // Retain document structure for bookmarks
                Bookmark = rootBookmark
            };

            // Save the workbook as PDF with the generated table of contents
            workbook.Save("WorkbookWithTOC.pdf", pdfOptions);
        }
    }
}
