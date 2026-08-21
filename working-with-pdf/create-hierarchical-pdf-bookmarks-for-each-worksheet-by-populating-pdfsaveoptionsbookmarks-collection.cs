// Title: Add hierarchical PDF bookmarks for each worksheet with Aspose.Cells PdfSaveOptions (C#)
// Description: Creates a workbook with multiple worksheets, places a title in cell A1 of each sheet, builds a root PdfBookmarkEntry and child entries for every worksheet, and saves the workbook as a PDF using PdfSaveOptions.Bookmark and ExportDocumentStructure to generate a navigable bookmark tree.
// Keywords: Aspose.Cells | PdfSaveOptions | PdfBookmarkEntry | C# PDF bookmarks | worksheet PDF outline | ExportDocumentStructure | Aspose.Cells PDF export | nested PDF bookmarks | C# workbook to PDF | Aspose.Cells example
// Common Searches: Aspose.Cells add PDF bookmarks C# | Create PDF outline from Excel worksheets | PdfSaveOptions Bookmark property example | Hierarchical PDF bookmarks Aspose.Cells | Export workbook to PDF with bookmarks | C# code for PDF bookmark tree Aspose.Cells
// Developer Intent: Create a PDF from a workbook where each worksheet appears as a bookmark under a single root entry.
// Use Cases: Navigate large multi‑sheet reports in PDF via a clickable outline | Generate a table of contents for financial statements exported from Excel | Provide end‑users quick access to specific sections of a PDF generated from a workbook | Automate PDF documentation with structured bookmarks for regulatory submissions
// AI Prompts: Write C# code that adds a root PDF bookmark and child bookmarks for every worksheet using Aspose.Cells. | Explain how ExportDocumentStructure affects bookmark creation in Aspose.Cells PDF export. | Show an example of building a nested PdfBookmarkEntry hierarchy for a workbook. | How to set PdfSaveOptions.Bookmark to include multiple levels of bookmarks in a PDF.

using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfBookmarks
{
    // Creates a workbook with multiple worksheets, places a title in cell A1 of each sheet, builds a root PdfBookmarkEntry and child entries for every worksheet, and saves the workbook as a PDF using PdfSaveOptions.Bookmark and ExportDocumentStructure to generate a navigable bookmark tree.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add several worksheets
            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Name = "Summary";
            workbook.Worksheets.Add("Sales");
            workbook.Worksheets.Add("Inventory");
            workbook.Worksheets.Add("Analysis");

            // Put a title in cell A1 of each sheet – this cell will be the bookmark destination
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.Cells["A1"].PutValue($"{sheet.Name} Content");
            }

            // Create the root bookmark entry
            PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
            {
                Text = "Workbook",
                Destination = workbook.Worksheets[0].Cells["A1"], // point to first sheet
                IsOpen = true,
                SubEntry = new ArrayList()
            };

            // Create a sub‑bookmark for each worksheet and add it to the root
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                PdfBookmarkEntry sheetBookmark = new PdfBookmarkEntry
                {
                    Text = sheet.Name,
                    Destination = sheet.Cells["A1"]
                };

                rootBookmark.SubEntry.Add(sheetBookmark);
            }

            // Configure PDF save options with the hierarchical bookmarks
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Bookmark = rootBookmark,
                ExportDocumentStructure = true   // ensures the bookmark tree is written to the PDF
            };

            // Save the workbook as a PDF file with the defined bookmarks
            workbook.Save("WorkbookBookmarks.pdf", pdfOptions);
        }
    }
}
