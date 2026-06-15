using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfTocDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add several worksheets
            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Name = "Summary";
            Worksheet sheet1 = workbook.Worksheets.Add("Data");
            Worksheet sheet2 = workbook.Worksheets.Add("Analysis");

            // Populate each sheet with some sample data
            PopulateSheet(workbook.Worksheets["Summary"], "Overview", 10);
            PopulateSheet(sheet1, "Item", 30);
            PopulateSheet(sheet2, "Result", 20);

            // Prepare image/print options for rendering (default options are sufficient)
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions();

            // Calculate page numbers for each worksheet
            int pageNumber = 1; // PDF pages start at 1
            PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
            {
                Text = "Table of Contents",
                IsOpen = true,
                SubEntry = new ArrayList()
            };

            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Render the worksheet to obtain its page count
                SheetRender sheetRender = new SheetRender(ws, renderOptions);
                int sheetPageCount = sheetRender.PageCount;

                // Create a bookmark entry for the worksheet.
                // The Text includes the sheet name and the starting page number.
                PdfBookmarkEntry entry = new PdfBookmarkEntry
                {
                    Text = $"{ws.Name} (Page {pageNumber})",
                    Destination = ws.Cells["A1"], // Link to the first cell of the sheet
                    IsOpen = false
                };

                // Add the entry to the root bookmark's sub‑entries
                rootBookmark.SubEntry.Add(entry);

                // Update the next starting page number
                pageNumber += sheetPageCount;
            }

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Export document structure (bookmarks) – optional because we set Bookmark manually
                ExportDocumentStructure = true,
                // Assign the constructed bookmark hierarchy
                Bookmark = rootBookmark
            };

            // Save the workbook as a PDF with the generated table of contents
            workbook.Save("WorkbookWithToc.pdf", pdfOptions);
        }

        // Helper method to fill a worksheet with simple data
        private static void PopulateSheet(Worksheet sheet, string header, int rows)
        {
            sheet.Cells["A1"].PutValue(header);
            for (int i = 0; i < rows; i++)
            {
                sheet.Cells[i + 1, 0].PutValue($"{header} {i + 1}");
                sheet.Cells[i + 1, 1].PutValue(i * 10);
            }
        }
    }
}