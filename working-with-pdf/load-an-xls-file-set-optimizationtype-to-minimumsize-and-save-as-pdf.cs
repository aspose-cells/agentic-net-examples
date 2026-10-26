// Title: Convert an XLS workbook to PDF using Aspose.Cells for .NET (C# example)
// Common Searches: how to convert an .xls file to PDF with Aspose.Cells in C# | Aspose.Cells C# example for saving workbook as PDF | C# code to load Excel .xls and export to PDF using Aspose.Cells | using PdfSaveOptions with Aspose.Cells to generate PDF from XLS | Aspose.Cells .NET convert legacy Excel format to PDF programmatically
// Tags: Aspose.Cells convert XLS to PDF | C# PdfSaveOptions usage Aspose.Cells | export Excel workbook as PDF .NET | Aspose.Cells workbook save as PDF example | load XLS file Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsExample
{
    // The sample loads an existing .xls file with Aspose.Cells, creates a PdfSaveOptions object, ensures the output directory exists, and saves the workbook as a PDF while handling possible exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xls";
            string outputPath = "output.pdf";

            try
            {
                // Verify that the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing XLS file
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options (default options are sufficient)
                PdfSaveOptions saveOptions = new PdfSaveOptions();

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as a PDF document
                workbook.Save(outputPath, saveOptions);

                Console.WriteLine($"PDF saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
