// Title: Export a localized chart to a high‑resolution PNG with Aspose.Cells for .NET
// Description: Shows how to build a workbook, add a column chart with Chinese category labels and title, configure ImageOrPrintOptions for 300 DPI PNG output, and export the chart using Chart.ToImage while optionally saving the workbook.
// Keywords: Aspose.Cells | C# chart export | high DPI PNG | localized chart title | Chinese labels | ImageOrPrintOptions | Chart.ToImage | .NET | 300 DPI | high resolution chart image
// Common Searches: Aspose.Cells export chart PNG 300 DPI | C# chart with Chinese title high resolution image | How to set DPI for chart export in Aspose.Cells | Chart.ToImage high‑resolution example .NET | Export localized chart as PNG using Aspose.Cells
// Developer Intent: Create a high‑resolution PNG of a chart that contains non‑Latin (e.g., Chinese) text.
// Use Cases: Generate a printable sales chart with Chinese labels at 300 DPI for marketing materials. | Save both the Excel workbook and a crisp PNG snapshot for documentation or reporting. | Swap ImageOrPrintOptions settings to produce JPEG or change DPI for different publishing needs.
// AI Prompts: Write C# code that exports an Aspose.Cells chart with a Japanese title to a 600 DPI PNG. | Explain how ImageOrPrintOptions controls image size, resolution, and format when using Chart.ToImage.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsChartExport
{
    // Shows how to build a workbook, add a column chart with Chinese category labels and title, configure ImageOrPrintOptions for 300 DPI PNG output, and export the chart using Chart.ToImage while optionally saving the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data (including localized text for the chart title)
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("苹果");   // Apple in Chinese
            worksheet.Cells["A3"].PutValue("橙子");   // Orange in Chinese
            worksheet.Cells["A4"].PutValue("香蕉");   // Banana in Chinese

            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(1200);
            worksheet.Cells["B3"].PutValue(800);
            worksheet.Cells["B4"].PutValue(1500);

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.SetChartDataRange("A1:B4", true);

            // Set a localized title to evaluate text clarity
            chart.Title.Text = "水果销售量"; // "Fruit Sales Volume" in Chinese

            // Configure high‑resolution image options (e.g., 300 DPI)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png,
                HorizontalResolution = 300,
                VerticalResolution = 300
            };

            // Export the chart to a high‑resolution PNG image
            string imagePath = "HighResChart.png";
            chart.ToImage(imagePath, imgOptions);

            // Optionally save the workbook for reference
            workbook.Save("ChartWorkbook.xlsx");

            Console.WriteLine($"Chart exported to high‑resolution image: {imagePath}");
        }
    }
}
