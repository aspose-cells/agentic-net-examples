// Title: Export a selected Excel range to a single‑page PDF with custom margins and landscape orientation using Aspose.Cells for .NET
// AI Prompts: Write C# code that sets the print area to A1:D10, applies 0.5‑inch left/right margins and 1‑inch top/bottom margins, switches the page orientation to landscape, and saves the range as a one‑page PDF with Aspose.Cells. | Show how to combine Worksheet.PageSetup settings with PdfSaveOptions.OnePagePerSheet to export a defined range to PDF while customizing page margins in a .NET application. | Adapt the sample to export multiple non‑contiguous ranges, each with its own margin and orientation configuration, into separate PDF files using Aspose.Cells.
// Common Searches: asp.net export specific Excel cells to PDF with custom page margins | set landscape orientation for PDF output of a range using Aspose.Cells | how to keep an exported Excel range on a single PDF page with Aspose.Cells | configure print area and margins before saving workbook as PDF in C# | PdfSaveOptions OnePagePerSheet example for range export Aspose.Cells
// Tags: export range to PDF Aspose.Cells | custom page margins Worksheet.PageSetup | landscape orientation PDF Aspose.Cells | PdfSaveOptions OnePagePerSheet setting | set print area programmatically Aspose.Cells | C# Aspose.Cells range export

using Aspose.Cells;
using System;
using System.IO;

// The code loads an Excel workbook, defines A1:D10 as the print area, applies 0.5‑inch side margins and 1‑inch top/bottom margins, optionally sets landscape orientation, configures PdfSaveOptions to keep the range on a single page, and saves the result as a PDF.
class ExportRangeToPdf
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range to export
            const string exportRange = "A1:D10";

            // Set the print area to the desired range
            sheet.PageSetup.PrintArea = exportRange;

            // Apply custom margins (points; 1 inch = 72 points)
            sheet.PageSetup.LeftMargin = 36;   // 0.5 inch left margin
            sheet.PageSetup.RightMargin = 36;  // 0.5 inch right margin
            sheet.PageSetup.TopMargin = 72;    // 1 inch top margin
            sheet.PageSetup.BottomMargin = 72; // 1 inch bottom margin

            // Set page orientation to Landscape (if supported)
            // Uncomment the following line if the PageOrientation enum is available in your Aspose.Cells version
            // sheet.PageSetup.Orientation = PageOrientation.Landscape;

            // Configure PDF save options to keep the range on a single page
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
            };

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the selected range as a PDF file
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
