// Title: Export an Excel workbook with slicers to a single-page PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file containing slicers, sets PdfSaveOptions to render each worksheet on one page, and saves the workbook as a PDF with Aspose.Cells. | Show how to check whether the source Excel file exists before conversion and throw a clear exception if it is missing. | Demonstrate configuring the PDF save options so that slicers stay on the same page when the workbook is exported.
// Common Searches: c# aspocells export workbook containing slicers to pdf single page | how to keep slicers on same page when saving Excel as PDF with Aspose.Cells | pdfsaveoptions single page per sheet example for slicer worksheets | validate input file existence before converting xlsx to pdf using Aspose.Cells
// Tags: Aspose.Cells PDF export with slicers | One-page-per-sheet PDF conversion | C# slicer preservation during PDF save | File existence check Aspose.Cells conversion | PdfSaveOptions configuration for single-page sheets

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// // Loads an .xlsx workbook that contains slicers, configures PdfSaveOptions to render each worksheet on a single page so slicers stay together, validates the input file, and saves the result as a PDF with proper exception handling.
class ExportWorkbookWithSlicers
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the existing workbook that contains slicers
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Export each worksheet as a single page so that all slicers stay on the same page
                OnePagePerSheet = true
                // Additional options (e.g., image resolution) can be set here if needed
            };

            // Save the workbook to PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully exported to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log or display the error details for troubleshooting
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
