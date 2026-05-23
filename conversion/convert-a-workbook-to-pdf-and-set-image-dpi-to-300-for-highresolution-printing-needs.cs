using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfExport
{
    public class WorkbookToPdfHighDpi
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                // Set the global DPI to 300 for high‑resolution rendering
                CellsHelper.DPI = 300;

                const string inputPath = "input.xlsx";
                const string outputPath = "output_high_dpi.pdf";

                // Verify that the source workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Resample images to 300 PPI and keep maximum JPEG quality (100)
                pdfOptions.SetImageResample(300, 100);

                // Set the print quality of the first worksheet to match the desired DPI
                workbook.Worksheets[0].PageSetup.PrintQuality = 300;

                // Save the workbook as a PDF with the specified options
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"PDF successfully saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}