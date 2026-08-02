// Title: Generate a PDF with a clickable Table of Contents from Excel sheets using Aspose.Cells for .NET
// Description: C# sample that builds a workbook, adds three worksheets filled with data, creates a PdfBookmarkEntry hierarchy (root "Table of Contents" with entries pointing to each sheet's A1 cell), enables ExportDocumentStructure, and saves the workbook as a PDF containing a clickable TOC for easy navigation.
// Keywords: Aspose.Cells PDF bookmarks | C# PDF table of contents | ExportDocumentStructure | PdfBookmarkEntry | Aspose.Cells save as PDF | Excel to PDF with TOC | Aspose.Cells .NET | PDF navigation bookmarks | generate PDF from workbook | Aspose.Cells PDFSaveOptions
// Common Searches: Aspose.Cells add PDF bookmarks | Create PDF table of contents from Excel using C# | How to export Excel workbook to PDF with TOC Aspose | PdfSaveOptions ExportDocumentStructure example | C# generate PDF with clickable TOC Aspose.Cells
// Developer Intent: Create a PDF from an Excel workbook that includes a clickable table of contents linking to the start of each worksheet.
// Use Cases: Automated multi‑sheet financial reports with a TOC for quick navigation | Exporting training manuals where each chapter is a worksheet and the PDF needs a navigation pane | Building accessible PDFs (PDF/UA) with document structure from Excel data | Generating product catalogs where each category is a sheet and the PDF includes a TOC
// AI Prompts: Show how to add page numbers to each entry in the PDF table of contents using Aspose.Cells. | Provide code to create nested bookmarks when worksheet names follow a hierarchical pattern. | Explain how to customize bookmark icons and colors in the generated PDF. | Give an example that places the table of contents on a separate first page with page numbers. | Convert this C# example to Aspose.Cells for Java while preserving the TOC functionality.

using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// C# sample that builds a workbook, adds three worksheets filled with data, creates a PdfBookmarkEntry hierarchy (root "Table of Contents" with entries pointing to each sheet's A1 cell), enables ExportDocumentStructure, and saves the workbook as a PDF containing a clickable TOC for easy navigation.
class GeneratePdfWithToc
{
    static void Main()
    {
        try
        {
            // Create a new workbook (contains one default worksheet)
            Workbook workbook = new Workbook();

            // Ensure we have three worksheets
            for (int i = 1; i < 3; i++)
            {
                workbook.Worksheets.Add();
            }

            // Populate each worksheet with sample data
            for (int i = 0; i < 3; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];
                sheet.Name = $"Sheet{i + 1}";
                sheet.Cells["A1"].PutValue($"Start of {sheet.Name}");

                // Fill rows to potentially span multiple PDF pages
                for (int row = 0; row < 100; row++)
                {
                    sheet.Cells[row, 0].PutValue($"Row {row + 1} in {sheet.Name}");
                }
            }

            // Create the root PDF bookmark (acts as the Table of Contents)
            PdfBookmarkEntry tocRoot = new PdfBookmarkEntry
            {
                Text = "Table of Contents",
                IsOpen = true,
                SubEntry = new ArrayList()
            };

            // Add a bookmark entry for each worksheet; destination is cell A1
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                PdfBookmarkEntry entry = new PdfBookmarkEntry
                {
                    Text = sheet.Name,
                    Destination = sheet.Cells["A1"]
                };
                tocRoot.SubEntry.Add(entry);
            }

            // Configure PDF save options to include document structure and bookmarks
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true,
                Bookmark = tocRoot
            };

            // Save the workbook as a PDF with the generated Table of Contents
            workbook.Save("WorkbookWithTOC.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
