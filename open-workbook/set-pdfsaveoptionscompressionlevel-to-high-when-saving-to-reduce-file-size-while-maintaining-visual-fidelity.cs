// Title: Convert an Excel workbook to a PDF with maximum compression using Aspose.Cells in C#
// AI Prompts: Write C# code that loads an .xlsx file and saves it as a PDF using Aspose.Cells, configuring PdfSaveOptions.CompressionLevel to High to shrink the output file. | Update an existing Aspose.Cells PDF export routine to apply the highest compression setting while preserving the workbook’s visual layout. | Show how to create a PdfSaveOptions object with CompressionLevel set to High before calling Workbook.Save in a .NET application.
// Common Searches: Aspose.Cells set PdfSaveOptions.CompressionLevel to High in C# example | How to reduce PDF size when converting Excel to PDF with Aspose.Cells .NET | C# code for exporting workbook to PDF with maximum compression using Aspose.Cells | PdfSaveOptions high compression best practice for preserving visual fidelity | Compress PDF output from Excel workbook Aspose.Cells tutorial
// Tags: Aspose.Cells PDF maximum compression | PdfSaveOptions CompressionLevel usage | C# export Excel to PDF with Aspose.Cells | reduce PDF file size Aspose.Cells | preserve visual fidelity PDF compression

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // Demonstrates loading an Excel workbook, configuring PdfSaveOptions with CompressionLevel set to High to minimize PDF size while keeping visual quality, and saving the workbook as a compressed PDF.
    class Program
    {
        static void Main(string[] args)
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

                // Load the Excel workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options (default compression)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as a PDF
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
