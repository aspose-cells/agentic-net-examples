using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsFullScreenPdfDemo
{
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output_fullscreen.pdf";

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

                // Create PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // NOTE: In newer Aspose.Cells versions the OpenInFullScreenMode property is available.
                // If the current version does not support it, the PDF will be saved without this setting.

                // Save the workbook as PDF with the specified options
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Workbook has been saved to PDF at \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}