// Title: Export an Excel workbook that contains XML maps to a single‑page PDF while keeping the visual layout using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with XML maps, configures PdfSaveOptions.OnePagePerSheet, and saves it as a PDF using Aspose.Cells. | Show how to add pre‑save validation that checks for the source Excel file, creates the destination folder if it does not exist, and handles exceptions during the PDF export. | Demonstrate configuring Aspose.Cells PDF conversion to retain the appearance of mapped cells and their data in the generated PDF.
// Common Searches: how to convert an Excel file that uses XML maps to PDF with Aspose.Cells preserving layout | Aspose.Cells C# export workbook to PDF one page per sheet example | C# verify input Excel exists and create output folder before saving PDF using Aspose.Cells | preserve cell formatting and XML mapping when saving Excel as PDF in .NET | PdfSaveOptions settings for keeping visual representation of mapped cells in Aspose.Cells
// Tags: Aspose.Cells PDF conversion with XML maps | single-page PDF export using Aspose.Cells | input file validation for Aspose.Cells workbook load | auto-create output folder for PDF save Aspose.Cells | maintain cell layout when converting Excel to PDF

using System;
using System.IO;
using Aspose.Cells;

// The program checks that the source Excel file exists, loads it into an Aspose.Cells Workbook, configures PdfSaveOptions to place each worksheet on a single PDF page, ensures the output directory is created, and saves the workbook as a PDF while preserving the visual representation of XML‑mapped cells and data.
class ExportToPdf
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
                // Compliance property omitted because it may not be available in all versions
            };

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
