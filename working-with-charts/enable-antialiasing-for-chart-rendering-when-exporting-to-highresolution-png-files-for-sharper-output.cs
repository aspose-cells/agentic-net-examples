using System;
using System.Drawing.Imaging;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    public class ChartAntiAliasingDemo
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["A3"].PutValue("Orange");
                worksheet.Cells["A4"].PutValue("Banana");

                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(1200);
                worksheet.Cells["B3"].PutValue(800);
                worksheet.Cells["B4"].PutValue(1500);

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.SetChartDataRange("A1:B4", true);

                // Configure image rendering options (high‑resolution PNG)
                ImageOrPrintOptions options = new ImageOrPrintOptions
                {
                    HorizontalResolution = 300, // High horizontal DPI
                    VerticalResolution = 300    // High vertical DPI
                    // Anti‑aliasing is enabled by default; ImageFormat is inferred from file extension
                };

                // Export the chart to a high‑resolution PNG
                string chartPath = "HighResChart.png";
                chart.ToImage(chartPath, options);

                // Save the workbook (optional)
                string workbookPath = "ChartWorkbook.xlsx";
                workbook.Save(workbookPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}