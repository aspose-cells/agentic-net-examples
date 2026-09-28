// Title: How to use Aspose.Cells PdfSaveOptions.FitAllColumnsOnOnePage = true in C# to fit every worksheet column onto a single PDF page
// AI Prompts: Write C# code that loads an Excel workbook and saves it as a PDF with all columns forced onto one page by setting PdfSaveOptions.FitAllColumnsOnOnePage to true. | Show an example of configuring Aspose.Cells PdfSaveOptions to fit all worksheet columns on a single PDF page in a .NET application. | Demonstrate error‑handled conversion of an .xlsx file to a single‑page PDF using Aspose.Cells with FitAllColumnsOnOnePage enabled.
// Common Searches: asp.net aspose.cells fit all columns on one pdf page c# | PdfSaveOptions FitAllColumnsOnOnePage property example | export excel to single page pdf with all columns using aspose cells | c# aspose cells pdf export column scaling single page | how to set FitAllColumnsOnOnePage true in Aspose.Cells
// Tags: PdfSaveOptions FitAllColumnsOnOnePage property | Aspose.Cells single-page PDF export | C# Excel to PDF column fitting | FitAllColumnsOnOnePage vs OnePagePerSheet | Aspose.Cells PDF column scaling

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfExport
{
    // The example loads an existing .xlsx workbook, creates a PdfSaveOptions object with FitAllColumnsOnOnePage set to true (ensuring all worksheet columns are compressed onto a single PDF page), and saves the workbook as a PDF. It also checks for the input file's existence and catches exceptions to report errors.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options to fit the entire sheet on a single page
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OnePagePerSheet = true
                };

                // Save the workbook as a PDF using the configured options
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
