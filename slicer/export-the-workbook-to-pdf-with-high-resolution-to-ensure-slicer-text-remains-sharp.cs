using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsHighResPdfExport
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Set a high DPI for rendering to improve image and shape quality.
                CellsHelper.DPI = 300;

                // Create a new workbook and obtain the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data.
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Fruits");
                sheet.Cells["A3"].PutValue("Vegetables");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(50);
                sheet.Cells["B3"].PutValue(30);

                // Add a simple column chart.
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B3", true);
                chart.NSeries.CategoryData = "A2:A3";
                chart.Title.Text = "Sample Chart";

                // Configure PDF save options for high quality.
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OptimizationType = PdfOptimizationType.Standard,
                    ExportDocumentStructure = true,
                    DefaultFont = "Arial",
                    EmbedStandardWindowsFonts = true
                };
                pdfOptions.SetImageResample(300, 90); // Resample images to 300 DPI with JPEG quality 90.

                // Define output path and ensure its directory exists.
                string outputPath = "HighResolutionOutput.pdf";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as a PDF.
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Workbook successfully saved to PDF with high resolution at: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}