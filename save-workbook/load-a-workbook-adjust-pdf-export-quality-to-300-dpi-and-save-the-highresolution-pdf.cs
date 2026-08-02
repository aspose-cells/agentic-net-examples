using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output_high_res.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Set the global DPI to 300 for high‑resolution rendering
            CellsHelper.DPI = 300;

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Resample images to 300 PPI and use maximum JPEG quality (100)
            pdfOptions.SetImageResample(300, 100);

            // (Optional) Use standard optimization for best print quality
            // pdfOptions.OptimizationType = PdfOptimizationOptions.Standard; // Uncomment if enum is available in your version

            // Save the workbook as a high‑resolution PDF
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}