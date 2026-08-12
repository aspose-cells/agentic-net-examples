// Title: C# – Export Excel to PDF with Worksheet‑Named Bookmarks using Aspose.Cells
// Description: The sample creates a hidden root entry, then adds a PdfBookmarkEntry for each worksheet where the entry text matches the sheet name and the destination points to cell A1. The bookmark hierarchy is attached to PdfSaveOptions (ExportDocumentStructure enabled) and the workbook is saved as a PDF that provides clickable navigation to every sheet.
// Keywords: Aspose.Cells | C# | .NET | PdfBookmarkEntry | PdfSaveOptions | Excel to PDF | worksheet bookmarks | PDF navigation | ExportDocumentStructure | Aspose.Cells example
// Common Searches: Aspose.Cells add PDF bookmarks per worksheet | C# export Excel workbook to PDF with bookmarks | How to set worksheet name as PDF bookmark title | Create hidden root bookmark Aspose.Cells | Enable ExportDocumentStructure PDF Aspose.Cells
// Developer Intent: Generate a PDF from an Excel workbook where each worksheet appears as a clickable bookmark titled with the sheet’s name.
// Use Cases: Distribute multi‑sheet reports that users can jump to directly from the PDF outline. | Automate generation of regulatory filings where each Excel tab becomes a separate PDF section. | Build printable catalogs where each product category is a worksheet and the PDF includes a navigation pane.
// AI Prompts: Modify the code to give the root bookmark a visible title such as "Workbook Index". | Add page‑number labels to each bookmarked section while exporting to PDF. | Group worksheets under parent bookmarks (e.g., "Q1", "Q2") while keeping individual sheet entries.

using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample creates a hidden root entry, then adds a PdfBookmarkEntry for each worksheet where the entry text matches the sheet name and the destination points to cell A1. The bookmark hierarchy is attached to PdfSaveOptions (ExportDocumentStructure enabled) and the workbook is saved as a PDF that provides clickable navigation to every sheet.
class PdfBookmarksExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Prepare worksheets with sample data
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "FirstSheet";
        sheet1.Cells["A1"].PutValue("Content of First Sheet");

        Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");
        sheet2.Cells["A1"].PutValue("Content of Second Sheet");

        Worksheet sheet3 = workbook.Worksheets.Add("ThirdSheet");
        sheet3.Cells["A1"].PutValue("Content of Third Sheet");

        // Create a hidden root bookmark entry
        PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
        {
            Text = "",               // Empty text hides the root entry
            SubEntry = new ArrayList(),
            IsOpen = true
        };

        // Add a bookmark for each worksheet using its name as the title
        foreach (Worksheet ws in workbook.Worksheets)
        {
            PdfBookmarkEntry entry = new PdfBookmarkEntry
            {
                Text = ws.Name,               // Bookmark title
                Destination = ws.Cells["A1"]  // Destination cell
            };
            rootBookmark.SubEntry.Add(entry);
        }

        // Configure PDF save options with the bookmark hierarchy
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Bookmark = rootBookmark,
            ExportDocumentStructure = true
        };

        // Save the workbook as a PDF file with bookmarks
        workbook.Save("WorksheetsBookmarks.pdf", saveOptions);
    }
}
