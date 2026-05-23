using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    public class WorkbookToPdfGrayscaleDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Item");
                sheet.Cells["B1"].PutValue("Quantity");
                sheet.Cells["A2"].PutValue("Apples");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("Bananas");
                sheet.Cells["B3"].PutValue(20);

                // Enable black‑and‑white (grayscale) rendering for each worksheet
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    ws.PageSetup.BlackAndWhite = true;
                }

                // Configure PDF save options for minimum file size
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OptimizationType = PdfOptimizationType.MinimumSize,
                    EmbedStandardWindowsFonts = true
                };

                // Define output path and ensure the directory exists
                string outputPath = "GrayscaleOutput.pdf";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as a PDF using the configured options
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook saved to PDF with grayscale rendering: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            WorkbookToPdfGrayscaleDemo.Run();
        }
    }
}