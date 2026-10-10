// Title: Generate a PDF from an Excel file with Aspose.Cells .NET while keeping original column widths and using MinimumSize optimization
// AI Prompts: Provide C# code that loads an .xlsx workbook, sets PdfSaveOptions.OptimizationType to PdfOptimizationType.MinimumSize, ensures column widths are not altered, and saves the workbook as a PDF. | Demonstrate how to configure Aspose.Cells PdfSaveOptions to produce the smallest possible PDF without affecting the worksheet column layout in a .NET application.
// Common Searches: Aspose.Cells .NET convert Excel to PDF keep column widths | PdfSaveOptions MinimumSize keep column layout C# | how to reduce PDF size from Excel using Aspose.Cells while preserving column dimensions
// Tags: Aspose.Cells PDF minimum size optimization | retain original column sizes Aspose.Cells | C# PdfSaveOptions column layout preservation | reduce PDF output size Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfExample
{
    // // Loads an Excel workbook, applies PdfSaveOptions with OptimizationType = MinimumSize to shrink the PDF while preserving the worksheet's column widths, and saves the result as a PDF.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook from the specified file
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Minimize the PDF file size
                    OptimizationType = PdfOptimizationType.MinimumSize
                };

                // Save the workbook as a PDF using the configured options
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
