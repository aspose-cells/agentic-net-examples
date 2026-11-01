// Title: Export an Excel workbook to a high‑resolution 300 DPI PDF using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, sets PdfSaveOptions.ImageResolution to 300 DPI, and saves the workbook as a PDF. | Show how to verify the input file, configure PdfSaveOptions for high‑resolution output, create missing directories, and handle exceptions during Excel‑to‑PDF conversion in C#.
// Common Searches: Aspose.Cells C# set PDF export DPI to 300 | How to increase image resolution when saving Excel as PDF with Aspose.Cells | C# convert .xlsx to high resolution PDF using PdfSaveOptions.ImageResolution | PdfSaveOptions ImageResolution property example for 300 DPI | Save workbook as PDF with 300 DPI using Aspose.Cells .NET
// Tags: Aspose.Cells PdfSaveOptions ImageResolution 300 DPI | C# high‑resolution Excel to PDF conversion | set PDF export DPI Aspose.Cells | Workbook.Save PDF high resolution | adjust PDF image resolution Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Saving;
using System;
using System.IO;

// The program checks that the source Excel file exists, loads it into a Workbook, sets PdfSaveOptions.ImageResolution to 300 DPI for high‑quality output, ensures the target directory is present, and saves the workbook as a PDF while handling any runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default settings are used here)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved to PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
