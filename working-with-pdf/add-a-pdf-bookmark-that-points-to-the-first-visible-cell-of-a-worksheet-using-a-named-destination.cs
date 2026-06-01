using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfBookmarkDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add some sample content – the first visible cell will be A1
            worksheet.Cells["A1"].PutValue("First visible cell content");

            // Create a PDF bookmark entry that points to the first visible cell (A1)
            PdfBookmarkEntry bookmark = new PdfBookmarkEntry
            {
                Text = "First Visible Cell",          // Title of the bookmark
                Destination = worksheet.Cells["A1"], // Cell the bookmark links to
                DestinationName = "FirstCellDest",   // Named destination
                IsOpen = true                        // Expand the bookmark in the PDF outline
            };

            // Configure PDF save options with the bookmark
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Bookmark = bookmark,
                ExportDocumentStructure = true // Ensure document structure (bookmarks) is exported
            };

            // Save the workbook as a PDF with the configured bookmark
            workbook.Save("FirstVisibleCellBookmark.pdf", pdfOptions);
        }
    }
}