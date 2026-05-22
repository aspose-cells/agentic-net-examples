using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];
            chart.SetChartDataRange("A1:B4", true);

            // Set legend position to the top‑right corner of the plot area
            chart.Legend.Position = LegendPositionType.Corner;

            // Recalculate the chart so that layout properties are up‑to‑date
            chart.Calculate();

            // Verify legend coordinates (ratio to chart area and pixel values)
            double legendXRatio = chart.Legend.XRatioToChart;   // 0‑1 range
            double legendYRatio = chart.Legend.YRatioToChart;
            int legendXPixel = chart.Legend.XPixel;            // pixel coordinates
            int legendYPixel = chart.Legend.YPixel;

            Console.WriteLine($"Legend X Ratio: {legendXRatio}");
            Console.WriteLine($"Legend Y Ratio: {legendYRatio}");
            Console.WriteLine($"Legend X Pixel: {legendXPixel}");
            Console.WriteLine($"Legend Y Pixel: {legendYPixel}");

            // Save the workbook
            workbook.Save("LegendTopRightCorner.xlsx");
        }
    }
}