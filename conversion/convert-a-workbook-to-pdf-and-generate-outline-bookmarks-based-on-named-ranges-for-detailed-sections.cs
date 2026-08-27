// Title: Convert an Excel workbook to PDF with worksheet outline bookmarks using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, enables ExportDocumentStructure in PdfSaveOptions, builds a PdfBookmarkEntry hierarchy for each worksheet pointing to cell A1, and saves the workbook as a PDF. | Show how to create a root PdfBookmarkEntry and add child entries for every worksheet to generate PDF outline bookmarks in Aspose.Cells. | Adapt the example to generate PDF bookmarks from named ranges instead of worksheet names while converting an Excel file to PDF.
// Common Searches: Aspose.Cells C# export Excel to PDF with outline bookmarks for each sheet | How to add PDF bookmarks when converting a workbook to PDF using Aspose.Cells .NET | PdfSaveOptions ExportDocumentStructure property usage example in C# | Create hierarchical PDF bookmarks from Excel worksheets with Aspose.Cells | C# generate PDF outline navigation from named ranges using Aspose.Cells
// Tags: Aspose.Cells export workbook to PDF with outline bookmarks | PdfSaveOptions ExportDocumentStructure property | PdfBookmarkEntry worksheet bookmark hierarchy | C# convert Excel to PDF with PDF outline navigation | Aspose.Cells generate PDF bookmarks from named ranges

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfBookmarkDemo
{
    // The program verifies the presence of input.xlsx, loads it with Aspose.Cells, configures PdfSaveOptions to export document structure, builds a root PdfBookmarkEntry containing a child entry for each worksheet that points to cell A1, assigns this bookmark hierarchy to the save options, and saves the workbook as output.pdf.
    class Program
    {
        static void Main()
        {
            try
            {
                // Verify that the input workbook exists
                const string inputFile = "input.xlsx";
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Error: Input file '{inputFile}' not found.");
                    return;
                }

                // Load the workbook (lifecycle rule: load)
                Workbook workbook = new Workbook(inputFile);

                // Create PDF save options (lifecycle rule: create)
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Export document structure for better outline navigation
                    ExportDocumentStructure = true
                };

                // Root bookmark entry
                PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
                {
                    Text = "Workbook Sections",
                    IsOpen = true,
                    SubEntry = new ArrayList()
                };

                // Iterate through all worksheets and create a bookmark for each sheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Destination cell is the first cell (A1) of the worksheet
                    Cell destinationCell = sheet.Cells[0, 0];

                    // Create a bookmark entry for the worksheet
                    PdfBookmarkEntry entry = new PdfBookmarkEntry
                    {
                        Text = sheet.Name,
                        Destination = destinationCell
                    };

                    // Add the entry to the root's sub‑entries
                    rootBookmark.SubEntry.Add(entry);
                }

                // Assign the constructed bookmark hierarchy to the PDF options
                pdfOptions.Bookmark = rootBookmark;

                // Save the workbook as PDF with the bookmark options (lifecycle rule: save)
                const string outputFile = "output.pdf";
                workbook.Save(outputFile, pdfOptions);
                Console.WriteLine($"PDF saved successfully to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
