// Title: Create PDF from Excel with Worksheet Bookmarks using Aspose.Cells for .NET
// Description: This example shows how to load or create an Excel workbook, add worksheets if needed, write a value to cell A1 of each sheet, build a root PdfBookmarkEntry and sub‑bookmarks named after each worksheet, enable ExportDocumentStructure in PdfSaveOptions, and save the workbook as a PDF that contains a navigable outline reflecting the sheet names.
// Keywords: Aspose.Cells | C# | .NET | Excel to PDF | PDF bookmarks | PdfBookmarkEntry | ExportDocumentStructure | worksheet outline | PDF navigation | document structure
// Common Searches: Aspose.Cells add PDF bookmarks from Excel sheets | C# save workbook as PDF with outline | Export Excel workbook to PDF with navigation bookmarks | PdfSaveOptions ExportDocumentStructure example | Create PDF bookmark hierarchy using Aspose.Cells
// Developer Intent: Generate a PDF from an Excel workbook and embed outline bookmarks for each worksheet to enable quick navigation.
// Use Cases: Produce a multi‑section PDF report where each Excel sheet appears as a clickable bookmark. | Automate the creation of user manuals that preserve the workbook’s tab structure in PDF form. | Provide end‑users with a searchable PDF that mirrors the original Excel file’s organization.
// AI Prompts: How can I set custom colors or styles for PdfBookmarkEntry bookmarks in Aspose.Cells? | Show me how to create nested PDF bookmarks for grouped worksheets using Aspose.Cells. | Provide code to save the PDF to a MemoryStream while keeping the bookmark hierarchy intact.

using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// This example shows how to load or create an Excel workbook, add worksheets if needed, write a value to cell A1 of each sheet, build a root PdfBookmarkEntry and sub‑bookmarks named after each worksheet, enable ExportDocumentStructure in PdfSaveOptions, and save the workbook as a PDF that contains a navigable outline reflecting the sheet names.
class WorkbookToPdfWithBookmarks
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // new Workbook("input.xlsx");

        // Ensure there are worksheets to work with
        if (workbook.Worksheets.Count == 0)
        {
            workbook.Worksheets.Add("Sheet1");
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");
        }

        // Populate each worksheet with a value at A1 (bookmark destination)
        foreach (Worksheet ws in workbook.Worksheets)
        {
            ws.Cells["A1"].PutValue($"{ws.Name} Content");
        }

        // Create the root bookmark entry
        PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
        {
            Text = "Workbook",
            Destination = workbook.Worksheets[0].Cells["A1"],
            IsOpen = true,
            SubEntry = new ArrayList()
        };

        // Add a sub‑bookmark for each worksheet using its name and A1 cell as destination
        foreach (Worksheet ws in workbook.Worksheets)
        {
            PdfBookmarkEntry sheetBookmark = new PdfBookmarkEntry
            {
                Text = ws.Name,
                Destination = ws.Cells["A1"]
            };
            rootBookmark.SubEntry.Add(sheetBookmark);
        }

        // Configure PDF save options to include the bookmark structure
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ExportDocumentStructure = true,
            Bookmark = rootBookmark
        };

        // Save the workbook as PDF with the generated outline bookmarks
        workbook.Save("WorkbookWithBookmarks.pdf", pdfOptions);
    }
}
