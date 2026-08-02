// Title: Add hierarchical worksheet bookmarks to a PDF with Aspose.Cells (C#)
// Description: Demonstrates how to create a root PdfBookmarkEntry, add child entries for each worksheet (pointing to cell A1), and save the workbook as a PDF with a nested bookmark outline using PdfSaveOptions.Bookmark and ExportDocumentStructure.
// Keywords: Aspose.Cells PDF bookmarks | PdfBookmarkEntry C# | PdfSaveOptions Bookmark property | hierarchical PDF outline | export Excel to PDF with bookmarks | Aspose.Cells workbook to PDF | C# Aspose.Cells example
// Common Searches: Aspose.Cells add PDF bookmarks for each worksheet | C# create hierarchical PDF bookmarks from Excel | PdfSaveOptions Bookmark example Aspose.Cells | Export workbook to PDF with outline navigation | How to set ExportDocumentStructure in Aspose.Cells
// Developer Intent: Generate a PDF from an Excel workbook where every worksheet appears as a navigable bookmark within a parent "Workbook" node.
// Use Cases: Produce a multi‑sheet financial report PDF with a top‑level bookmark and child entries for quick sheet navigation. | Create a PDF portfolio where each Excel worksheet is represented in the PDF bookmark tree for easy reference. | Automate PDF generation from Excel templates that require a structured outline for end‑user documentation.
// AI Prompts: Show C# code to add nested PDF bookmarks for each worksheet using Aspose.Cells. | Explain how to configure PdfSaveOptions.Bookmark with multiple levels and set ExportDocumentStructure. | Provide a step‑by‑step guide to create a root bookmark and child worksheet entries in Aspose.Cells.

using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfBookmarks
{
    // Demonstrates how to create a root PdfBookmarkEntry, add child entries for each worksheet (pointing to cell A1), and save the workbook as a PDF with a nested bookmark outline using PdfSaveOptions.Bookmark and ExportDocumentStructure.
    class Program
    {
        static void Main()
        {
            // Create a new workbook with three worksheets
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");

            // Put a marker value in each sheet – this cell will be the bookmark destination
            sheet1.Cells["A1"].PutValue("Content of Sheet1");
            sheet2.Cells["A1"].PutValue("Content of Sheet2");
            sheet3.Cells["A1"].PutValue("Content of Sheet3");

            // Create the root bookmark entry
            PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
            {
                Text = "Workbook",
                IsOpen = true,
                SubEntry = new ArrayList()   // initialize the collection for child entries
            };

            // Helper method to add a worksheet bookmark to the root
            void AddWorksheetBookmark(Worksheet ws)
            {
                PdfBookmarkEntry entry = new PdfBookmarkEntry
                {
                    Text = ws.Name,
                    Destination = ws.Cells["A1"]
                };
                rootBookmark.SubEntry.Add(entry);
            }

            // Add a bookmark for each worksheet
            AddWorksheetBookmark(sheet1);
            AddWorksheetBookmark(sheet2);
            AddWorksheetBookmark(sheet3);

            // Configure PDF save options with the hierarchical bookmark
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Bookmark = rootBookmark,
                ExportDocumentStructure = true   // ensure the document structure is exported
            };

            // Save the workbook as a PDF file with the defined bookmarks
            workbook.Save("WorkbookWithBookmarks.pdf", pdfOptions);
        }
    }
}
