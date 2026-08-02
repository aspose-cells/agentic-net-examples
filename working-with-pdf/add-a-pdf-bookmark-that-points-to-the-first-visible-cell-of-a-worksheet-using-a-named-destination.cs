// Title: Add a PDF bookmark to the first visible cell with Aspose.Cells for .NET
// Description: This example shows how to create a Workbook, locate the first non‑hidden cell, define a named destination, configure a PdfBookmarkEntry, enable document structure, and save the sheet as a PDF that opens directly to that cell.
// Keywords: Aspose.Cells PDF bookmark | first visible cell bookmark | named destination Aspose.Cells | PdfSaveOptions ExportDocumentStructure | C# Excel to PDF with bookmarks
// Common Searches: Aspose.Cells add PDF bookmark to a cell | how to find first visible cell in worksheet C# | save Excel as PDF with named destination bookmark | export workbook to PDF with document structure Aspose
// Developer Intent: Generate a PDF that includes a bookmark pointing to the worksheet's first visible cell using a named destination.
// Use Cases: Create navigable PDF reports where the bookmark jumps to the initial data row. | Combine multiple worksheets into one PDF, each with an expandable bookmark that opens at its first visible cell. | Produce PDFs with document structure so PDF viewers automatically display the bookmark pane.
// AI Prompts: Write C# code with Aspose.Cells to add a PDF bookmark that targets the first non‑hidden cell and save the workbook as PDF. | Explain how to locate the first visible cell in an Aspose.Cells worksheet and use it as a named destination in PdfBookmarkEntry. | Provide step‑by‑step instructions for configuring PdfSaveOptions with ExportDocumentStructure and a PdfBookmarkEntry that includes DestinationName.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfBookmarkExample
{
    // This example shows how to create a Workbook, locate the first non‑hidden cell, define a named destination, configure a PdfBookmarkEntry, enable document structure, and save the sheet as a PDF that opens directly to that cell.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some data so the worksheet has visible content
                sheet.Cells["A1"].PutValue("First visible cell");
                sheet.Cells["B2"].PutValue("Another cell");

                // Determine the first visible (non‑hidden) cell in the worksheet
                Cell firstVisible = GetFirstVisibleCell(sheet);

                // Create a PDF bookmark entry that points to the first visible cell
                PdfBookmarkEntry bookmark = new PdfBookmarkEntry
                {
                    Text = "Go to First Visible Cell",   // Title shown in PDF bookmarks pane
                    Destination = firstVisible,          // Cell the bookmark will navigate to
                    DestinationName = "FirstCellDest",   // Named destination for the cell
                    IsOpen = true                        // Expand this bookmark when PDF is opened
                };

                // Configure PDF save options with the bookmark
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    Bookmark = bookmark,
                    ExportDocumentStructure = true      // Ensure document structure (including bookmarks) is exported
                };

                // Define output PDF path
                string outputPath = "WorkbookWithBookmark.pdf";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as a PDF file with the configured bookmark
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        private static Cell GetFirstVisibleCell(Worksheet sheet)
        {
            // Determine the range that contains data
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            // If there is no data, fallback to A1
            if (maxRow < 0 || maxCol < 0)
                return sheet.Cells[0, 0];

            for (int row = 0; row <= maxRow; row++)
            {
                // Skip hidden rows
                if (sheet.Cells.Rows[row].IsHidden)
                    continue;

                for (int col = 0; col <= maxCol; col++)
                {
                    // Skip hidden columns
                    if (sheet.Cells.Columns[col].IsHidden)
                        continue;

                    Cell cell = sheet.Cells[row, col];

                    // Return the first non‑empty cell; if you prefer the first cell regardless of content, remove the null check
                    if (cell.Value != null)
                        return cell;
                }
            }

            // Fallback if all cells are hidden or empty
            return sheet.Cells[0, 0];
        }
    }
}
