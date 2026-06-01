using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    public class WorkbookToPdfLossless
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and add some sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sample Text");
                sheet.Cells["B1"].PutValue(12345);
                sheet.Cells["A2"].PutValue("Another Row");
                sheet.Cells["B2"].PutValue(67890);

                // Configure PDF save options to use lossless (Flate) compression
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Flate compression is lossless for images and other PDF content
                    PdfCompression = PdfCompressionCore.Flate,
                    // Use standard optimization to keep visual fidelity
                    OptimizationType = PdfOptimizationType.Standard
                };

                string outputPath = "output_lossless.pdf";

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as a PDF file with the specified options
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error during PDF conversion: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}