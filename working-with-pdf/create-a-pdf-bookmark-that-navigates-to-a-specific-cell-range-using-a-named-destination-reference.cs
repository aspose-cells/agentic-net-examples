// Title: Add a PDF bookmark to a named cell range with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, define a named range (MyRange) covering B2:C3, and attach a PdfBookmarkEntry that references the named destination. The bookmark is added to PdfSaveOptions and the workbook is saved as a PDF with a clickable bookmark that jumps to the specified range.
// Keywords: Aspose.Cells PDF bookmark | named destination Aspose.Cells | PdfBookmarkEntry C# | save workbook as PDF with bookmark | .NET Excel to PDF navigation | named range PDF link
// Common Searches: Aspose.Cells add PDF bookmark to named range | PdfBookmarkEntry DestinationName example | C# create PDF bookmark for specific cells | Aspose.Cells PDF navigation using named destinations | how to link PDF bookmark to Excel range
// Developer Intent: Create a PDF bookmark that points to a predefined cell range using a named destination.
// Use Cases: Generate a PDF report where a bookmark opens directly to a summary table defined as a named range. | Provide quick navigation in large spreadsheet PDFs by linking bookmarks to key data sections. | Add multiple bookmarks, each referencing a different named range, for fast access to various worksheet areas.
// AI Prompts: Show C# code to add several PDF bookmarks, each using a different named range, with Aspose.Cells. | Explain how the DestinationName property works for PDF bookmarks in Aspose.Cells and how to reference a named range. | Generate a sample that saves a workbook as PDF with a bookmark that opens at the first cell of a named range.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfBookmarkNamedDestination
{
    // Demonstrates how to create a workbook, define a named range (MyRange) covering B2:C3, and attach a PdfBookmarkEntry that references the named destination. The bookmark is added to PdfSaveOptions and the workbook is saved as a PDF with a clickable bookmark that jumps to the specified range.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Populate some cells that will be part of the named range
            sheet.Cells["B2"].PutValue("Start of Range");
            sheet.Cells["C3"].PutValue("End of Range");

            // Define a named range that covers B2:C3
            int nameIndex = workbook.Worksheets.Names.Add("MyRange");
            Name namedRange = workbook.Worksheets.Names[nameIndex];
            // The RefersTo formula must start with '='
            namedRange.RefersTo = "=Sheet1!$B$2:$C$3";

            // Create a PDF bookmark entry that uses the named destination
            PdfBookmarkEntry bookmark = new PdfBookmarkEntry
            {
                Text = "Bookmark to MyRange",
                // Destination can be any cell within the range; here we use the first cell
                Destination = sheet.Cells["B2"],
                // Set the named destination reference
                DestinationName = "MyRange",
                IsOpen = true
            };

            // Configure PDF save options with the bookmark
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Bookmark = bookmark
            };

            // Save the workbook as a PDF with the bookmark
            workbook.Save("BookmarkNamedDestination.pdf", pdfOptions);
        }
    }
}
