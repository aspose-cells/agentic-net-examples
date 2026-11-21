// Title: Convert an Excel workbook to PDF with each sheet on a separate page and all columns fitted using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export Excel to PDF with one page per worksheet | How to fit all columns on a single PDF page using Aspose.Cells PdfSaveOptions | Aspose.Cells PdfSaveOptions OnePagePerSheet and AllColumnsInOnePagePerSheet example | Create output folder programmatically before saving PDF with Aspose.Cells | Save Excel file as PDF with custom layout options in .NET
// Tags: Aspose.Cells PdfSaveOptions OnePagePerSheet | Aspose.Cells fit columns single PDF page | C# workbook.Save PDF with custom options | Aspose.Cells ensure output directory exists | Excel to PDF conversion Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file into a Workbook, configures PdfSaveOptions to place each worksheet on its own PDF page and to fit all columns onto a single page, creates the output directory if necessary, and then saves the workbook as a PDF using workbook.Save with the specified options.
class PdfExportExample
{
    static void Main()
    {
        // Paths for input Excel file and output PDF file
        string inputPath = "input.xlsx";
        string outputPath = "output.pdf";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The input file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook from the Excel file
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF export options
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
            {
                // Export each worksheet to a separate page
                OnePagePerSheet = true,

                // Fit all columns of a sheet onto one PDF page
                AllColumnsInOnePagePerSheet = true
            };

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as PDF using the configured options
            workbook.Save(outputPath, pdfSaveOptions);
            Console.WriteLine($"PDF successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
