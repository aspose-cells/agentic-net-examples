using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Charts;

namespace AsposeCellsHighResPdf
{
    public class ConvertChartToPdfHighResolution
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["A3"].PutValue("Orange");
            worksheet.Cells["A4"].PutValue("Banana");

            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(120);
            worksheet.Cells["B3"].PutValue(80);
            worksheet.Cells["B4"].PutValue(150);

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure PDF save options for high‑resolution images
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            pdfOptions.SetImageResample(300, 100); // 300 DPI, 100% JPEG quality
            pdfOptions.OptimizationType = PdfOptimizationType.Standard;

            // Define output file path
            string outputPath = "HighResolutionChart.pdf";

            // Save workbook to PDF
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook with chart saved to PDF at high image resolution: {outputPath}");
        }
    }
}