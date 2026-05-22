using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace ExportHighResChart
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["A4"].PutValue("Banana");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["B3"].PutValue(800);
            sheet.Cells["B4"].PutValue(1500);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.SetChartDataRange("A1:B4", true);

            // Configure high‑resolution image options
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png,          // PNG format
                HorizontalResolution = 300,         // 300 DPI horizontal
                VerticalResolution = 300            // 300 DPI vertical
            };

            // Export the chart to a high‑resolution PNG file
            string outputImagePath = "high_res_chart.png";
            chart.ToImage(outputImagePath, imgOptions);

            // Optionally save the workbook for reference
            workbook.Save("ChartWorkbook.xlsx");

            Console.WriteLine($"Chart exported to high‑resolution image: {outputImagePath}");
        }
    }
}