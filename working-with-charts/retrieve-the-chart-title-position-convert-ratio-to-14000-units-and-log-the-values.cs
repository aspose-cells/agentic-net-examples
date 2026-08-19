// Title: Retrieve Chart Title Position Ratios and Convert to 1/4000 Units – Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds a column chart with a title, reads the title's XRatioToChart and YRatioToChart (fraction of chart size), converts those ratios to integer values in 1/4000 units, logs the results, and saves the file.
// Keywords: Aspose.Cells | C# chart title position | XRatioToChart | YRatioToChart | convert ratio to 1/4000 units | Excel chart layout | retrieve chart title coordinates | .NET Excel automation
// Common Searches: Aspose.Cells get chart title XRatioToChart | How to convert chart title position to 1/4000 units in C# | Retrieve chart title YRatioToChart Aspose.Cells | Aspose.Cells chart title coordinates example | C# Aspose.Cells chart title placement
// Developer Intent: Read a chart title's X and Y ratios, transform them into 1/4000‑unit values, and output the numbers.
// Use Cases: Precisely align a chart title with other report elements by using absolute unit coordinates. | Compare title placements across multiple charts for consistent visual design. | Export title position data for documentation or automated layout validation.
// AI Prompts: Generate C# code with Aspose.Cells that reads a chart title's XRatioToChart and YRatioToChart, converts them to 1/4000 units, and prints the values. | Explain the meaning of XRatioToChart/YRatioToChart in Aspose.Cells and how to calculate absolute positions for Excel chart titles. | Provide a step‑by‑step example that retrieves chart title ratios, converts them, and saves the workbook while logging the results.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a column chart with a title, reads the title's XRatioToChart and YRatioToChart (fraction of chart size), converts those ratios to integer values in 1/4000 units, logs the results, and saves the file.
    public class RetrieveChartTitlePosition
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

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set the chart title
            chart.Title.Text = "Sample Chart Title";

            // Retrieve the title position ratios (fraction of chart width/height)
            double xRatio = chart.Title.XRatioToChart;
            double yRatio = chart.Title.YRatioToChart;

            // Convert ratios to 1/4000 units (as per documentation)
            int xInUnits = (int)Math.Round(xRatio * 4000);
            int yInUnits = (int)Math.Round(yRatio * 4000);

            // Output the values
            Console.WriteLine($"Title XRatioToChart (fraction): {xRatio}");
            Console.WriteLine($"Title YRatioToChart (fraction): {yRatio}");
            Console.WriteLine($"Title X position in 1/4000 units: {xInUnits}");
            Console.WriteLine($"Title Y position in 1/4000 units: {yInUnits}");

            // Save the workbook (optional, just to complete lifecycle)
            workbook.Save("RetrieveChartTitlePosition.xlsx");
        }
    }
}
