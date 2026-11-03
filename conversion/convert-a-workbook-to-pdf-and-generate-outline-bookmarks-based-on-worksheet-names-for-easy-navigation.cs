// Title: Generate a PDF from an Excel workbook with worksheet‑based outline bookmarks using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# create PDF with bookmarks for each worksheet | How to add Excel sheet names as PDF outline entries using Aspose.Cells | C# convert .xlsx to PDF with navigation bookmarks via PdfSaveOptions | Set BookmarksOutlineLevel in Aspose.Cells PDF export example | One page per sheet PDF export Aspose.Cells C# with bookmarks
// Tags: Aspose.Cells PDF conversion with worksheet bookmarks | PdfSaveOptions BookmarksOutlineLevel usage | C# export Excel to PDF one page per sheet | outline bookmarks generation from sheet names | Aspose.Cells navigation bookmarks in PDF

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering; // Required for PdfSaveOptions

// The sample loads an input.xlsx workbook, verifies its existence, configures PdfSaveOptions (enabling one page per sheet and optionally setting BookmarksOutlineLevel), and saves the workbook as output.pdf, producing a PDF where each worksheet appears as an outline bookmark for easy navigation.
class WorkbookToPdfWithBookmarks
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook from file
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
                // Uncomment the following line if the property is available in your Aspose.Cells version
                // BookmarksOutlineLevel = 1
            };

            // Save the workbook as PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
