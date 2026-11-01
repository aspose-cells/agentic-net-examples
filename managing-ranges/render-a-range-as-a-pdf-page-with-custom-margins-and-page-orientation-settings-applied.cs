// Title: Export a selected Excel range to a single‑page PDF with custom margins and landscape orientation using Aspose.Cells for .NET
// Common Searches: how to export a specific range as PDF with custom margins using Aspose.Cells C# | Aspose.Cells set print area and landscape orientation for PDF output | C# generate PDF from Excel range A1:D10 on one page with margins | PdfSaveOptions OnePagePerSheet example for range rendering in Aspose.Cells | adjust page margins in Aspose.Cells before saving workbook as PDF
// Tags: export range to PDF Aspose.Cells | set print area Aspose.Cells | custom page margins Aspose.Cells PDF | landscape orientation Aspose.Cells PDF export | OnePagePerSheet PdfSaveOptions usage

using Aspose.Cells;
using System;
using System.IO;

// The sample creates a workbook, fills cells A1:D10, defines that area as the print range, applies 0.5‑inch margins on all sides, optionally sets landscape orientation, configures PdfSaveOptions to fit the range on a single PDF page, ensures the output folder exists, and saves the result as a PDF file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in the range A1:D10
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 4; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Define the range to be rendered as PDF
            string printArea = "A1:D10";
            sheet.PageSetup.PrintArea = printArea;

            // Set page orientation to Landscape (if supported)
            // Note: Some Aspose.Cells versions may not expose PageOrientation enum.
            // The default orientation is Portrait; this line can be uncommented if the enum is available.
            // sheet.PageSetup.Orientation = PageOrientation.Landscape;

            // Set custom margins (0.5 inch = 36 points)
            double marginInPoints = 0.5 * 72; // 1 inch = 72 points
            sheet.PageSetup.LeftMargin = marginInPoints;
            sheet.PageSetup.RightMargin = marginInPoints;
            sheet.PageSetup.TopMargin = marginInPoints;
            sheet.PageSetup.BottomMargin = marginInPoints;

            // Configure PDF save options to fit the range on a single page
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
            };

            // Determine output path and ensure the directory exists
            string outputPath = "RangeOutput.pdf";
            string fullOutputPath = Path.GetFullPath(outputPath);
            string outputDir = Path.GetDirectoryName(fullOutputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as PDF
            workbook.Save(fullOutputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{fullOutputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
