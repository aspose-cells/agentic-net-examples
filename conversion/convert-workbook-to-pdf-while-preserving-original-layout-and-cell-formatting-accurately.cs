// Title: Convert an Excel workbook to PDF with original layout and cell formatting using Aspose.Cells for .NET (C#)
// Common Searches: c# convert excel workbook to pdf preserving cell formatting using aspose.cells | how to keep original column widths when saving excel as pdf with aspose.cells | aspnet core save xlsx as pdf/a-1b with aspose.cells | pdfsaveoptions onepagepersheet false example in c# | asp.net convert large excel file to multi-page pdf with aspose.cells
// Tags: Aspose.Cells PDF layout preservation | C# Excel to PDF/A-1b conversion | PdfSaveOptions disable OnePagePerSheet | retain column widths Aspose.Cells PDF | AllColumnsInOnePagePerSheet false setting | export workbook to multi-page PDF

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace WorkbookToPdfExample
{
    // The example loads an .xlsx file with Aspose.Cells, configures PdfSaveOptions (OnePagePerSheet = false, AllColumnsInOnePagePerSheet = false, PDF/A‑1b compliance) to retain original column widths, row heights, and cell formatting, ensures the output directory exists, and saves the workbook as a multi‑page PDF while handling missing files and runtime errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file
            string inputPath = "input.xlsx";

            // Path where the PDF will be saved
            string outputPath = "output.pdf";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                    return;
                }

                // Load the workbook from the specified file
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Keep the original column widths and row heights
                    OnePagePerSheet = false,
                    // Do not force all columns onto a single page
                    AllColumnsInOnePagePerSheet = false,
                    // Set PDF/A compliance if required
                    Compliance = PdfCompliance.PdfA1b
                };

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as a PDF using the configured options
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine("Workbook has been successfully converted to PDF.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
