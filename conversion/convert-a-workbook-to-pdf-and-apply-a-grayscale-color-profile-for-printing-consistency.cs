using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    public class WorkbookToPdfGrayscale
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Input workbook path
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Apply grayscale (black and white) printing setting to each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    sheet.PageSetup.BlackAndWhite = true; // Render sheet in black and white
                }

                // Configure PDF save options (default options are sufficient for grayscale output)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Output PDF path
                string outputPath = "output_grayscale.pdf";

                // Save the workbook as a PDF file with the grayscale settings applied
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Workbook has been converted to PDF with grayscale profile: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}