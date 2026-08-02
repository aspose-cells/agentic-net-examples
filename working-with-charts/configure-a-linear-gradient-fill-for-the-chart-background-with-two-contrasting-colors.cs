using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsGradientChartBackground
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // (Optional) Add some sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Item 1");
            sheet.Cells["A3"].PutValue("Item 2");
            sheet.Cells["A4"].PutValue("Item 3");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the fill format of the chart's plot area (background)
            FillFormat plotAreaFill = chart.PlotArea.Area.FillFormat;

            // Set the fill type to gradient to enable gradient properties
            plotAreaFill.FillType = FillType.Gradient;

            // Apply a linear two‑color gradient (e.g., black to white) horizontally
            plotAreaFill.SetTwoColorGradient(
                Color.Black,               // First (start) color
                Color.White,               // Second (end) color
                GradientStyleType.Horizontal, // Linear (horizontal) style
                1);                        // Variant (1‑4)

            // Save the workbook with the configured chart background
            workbook.Save("ChartWithLinearGradientBackground.xlsx");
        }
    }
}