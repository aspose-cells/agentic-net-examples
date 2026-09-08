// Title: How to convert an XLS file to PDF with MinimumSize optimization in C# using Aspose.Cells
// AI Prompts: Write C# code that loads an .xls workbook, sets PdfSaveOptions.OptimizationType = OptimizationType.MinimumSize, and saves it as a PDF with Aspose.Cells. | Show the steps to configure Aspose.Cells PDFSaveOptions for the smallest possible PDF size when converting a legacy Excel file in .NET.
// Common Searches: Aspose.Cells C# set PDF optimization to MinimumSize when converting XLS to PDF | How to reduce PDF file size using Aspose.Cells PDFSaveOptions in .NET | Convert legacy .xls workbook to PDF with minimum size using Aspose.Cells example
// Tags: Aspose.Cells PDFSaveOptions MinimumSize | C# XLS to PDF conversion Aspose.Cells | optimize PDF size Aspose.Cells | set PDF optimization type .NET | convert legacy Excel to PDF Aspose

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The program verifies that the input XLS file exists, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object with OptimizationType set to MinimumSize, and saves the workbook as a PDF. It reports success or any errors encountered during the process.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xls";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing XLS file
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options (default settings)
                PdfSaveOptions saveOptions = new PdfSaveOptions();

                string outputPath = "output.pdf";

                // Save the workbook as a PDF file using the configured options
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"PDF saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
