// Title: Set PdfSaveOptions.OnePagePerSheet to export each Excel worksheet as a single PDF page using Aspose.Cells for .NET
// Common Searches: how to export each Excel worksheet to a single PDF page with Aspose.Cells C# | Aspose.Cells PdfSaveOptions OnePagePerSheet example for .NET | save workbook as PDF with one page per sheet using Aspose.Cells in C# | C# code to create output folder before saving PDF with Aspose.Cells | handling missing Excel input file when converting to PDF with Aspose.Cells
// Tags: PdfSaveOptions OnePagePerSheet configuration | Aspose.Cells workbook to PDF conversion | single-page-per-sheet PDF export | ensure output directory exists before saving PDF | validate input file existence for Excel to PDF conversion

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The example checks that the source Excel file exists, loads it into a Workbook, configures PdfSaveOptions with OnePagePerSheet=true to place each worksheet on a single PDF page, creates the output directory if needed, and saves the workbook as a PDF while handling any runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true // Export each worksheet on a single page
                // Compliance property omitted for compatibility with older Aspose.Cells versions
            };

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
