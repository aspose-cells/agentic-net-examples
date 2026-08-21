// Title: Export Excel Workbook to PDF with Worksheet Bookmarks using Aspose.Cells for .NET (C#)
// Description: Loads an .xlsx file, builds a root PDF bookmark, adds a child bookmark for every worksheet (using the sheet name and cell A1 as the destination), enables document structure via PdfSaveOptions, and saves the workbook as a PDF that contains a clickable outline.
// Keywords: Aspose.Cells PDF bookmarks | C# export Excel to PDF | PdfBookmarkEntry | ExportDocumentStructure | worksheet bookmarks PDF | Aspose.Cells PDF outline | convert workbook to PDF | Aspose.Cells .NET
// Common Searches: Aspose.Cells add PDF bookmarks from Excel sheets | C# create PDF outline when saving workbook | Export Excel to PDF with bookmarks Aspose | PdfSaveOptions ExportDocumentStructure example | How to generate PDF bookmarks for each worksheet using Aspose.Cells
// Developer Intent: Create a PDF from an Excel workbook that includes a hierarchical bookmark for each worksheet.
// Use Cases: Generate a single PDF report where each worksheet appears as a clickable bookmark for fast navigation. | Automate production of PDF documentation that mirrors a multi‑sheet workbook, preserving the sheet hierarchy as a bookmark tree. | Export financial models or dashboards to PDF while keeping sheet names as navigable sections for end users.
// AI Prompts: Write C# code with Aspose.Cells to export an Excel file to PDF and add a bookmark for each worksheet, ensuring the PDF includes document structure. | Explain how ExportDocumentStructure and PdfBookmarkEntry work together in Aspose.Cells to create PDF outline bookmarks. | Modify the sample to set custom page numbers as bookmark destinations instead of using cell A1.

using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an .xlsx file, builds a root PDF bookmark, adds a child bookmark for every worksheet (using the sheet name and cell A1 as the destination), enables document structure via PdfSaveOptions, and saves the workbook as a PDF that contains a clickable outline.
class WorkbookToPdfWithBookmarks
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Create the root bookmark entry
        PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
        {
            Text = "Workbook",
            IsOpen = true
        };

        // Collection for child bookmarks (one per worksheet)
        ArrayList subEntries = new ArrayList();

        // Generate a bookmark for each worksheet using its name as the title
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Destination cell for the bookmark (A1 of each sheet)
            Cell destCell = sheet.Cells["A1"];

            // Ensure the destination cell contains something (optional)
            if (destCell.Value == null)
                destCell.PutValue(sheet.Name);

            // Create a bookmark entry for the current sheet
            PdfBookmarkEntry entry = new PdfBookmarkEntry
            {
                Text = sheet.Name,
                Destination = destCell
            };

            subEntries.Add(entry);
        }

        // Attach the child entries to the root bookmark
        rootBookmark.SubEntry = subEntries;

        // Configure PDF save options to include document structure and the bookmarks
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ExportDocumentStructure = true,
            Bookmark = rootBookmark
        };

        // Save the workbook as PDF with the defined options
        workbook.Save("output.pdf", pdfOptions);
    }
}
