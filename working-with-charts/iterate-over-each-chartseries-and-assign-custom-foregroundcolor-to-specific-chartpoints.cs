// Title: Set custom foreground colors for each ChartPoint in an Aspose.Cells column chart (C#)
// Description: The sample creates a workbook, adds a column chart with two series, then iterates through every series and its points. For each ChartPoint it assigns a specific ForegroundColor to the Area and sets the Formatting to Custom, ensuring the colors appear in the saved Excel file.
// Keywords: Aspose.Cells | C# | ChartPoint color | foreground color | custom chart point formatting | column chart | Excel chart customization | Series.Points | FormattingType.Custom | set chart point color programmatically
// Common Searches: Aspose.Cells change color of individual chart points | C# set chart point foreground color Aspose | how to customize column chart colors in Aspose.Cells | iterate chart series points Aspose.Cells | apply custom formatting to chart points .NET
// Developer Intent: Programmatically assign custom foreground colors to each point of all series in an Aspose.Cells column chart using C#.
// Use Cases: Highlight alternating columns within a series to improve readability. | Distinguish multiple series by applying different color patterns to their points. | Create a heat‑map style column chart by coloring points based on value ranges. | Generate branded reports where specific data points need brand‑specific colors.
// AI Prompts: Generate C# code with Aspose.Cells that colors chart points based on numeric thresholds. | Show how to color every third point red in a line chart using Aspose.Cells ChartPoint formatting. | Explain how to reset chart point formatting to default after applying custom colors.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // The sample creates a workbook, adds a column chart with two series, then iterates through every series and its points. For each ChartPoint it assigns a specific ForegroundColor to the Area and sets the Formatting to Custom, ensuring the colors appear in the saved Excel file.
    public class ChartSeriesPointColorDemo
    {
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
            sheet.Cells["A5"].PutValue("D");

            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["B5"].PutValue(40);

            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);
            sheet.Cells["C5"].PutValue(45);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for both series
            chart.NSeries.Add("B2:C5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Iterate over each series in the chart
            for (int s = 0; s < chart.NSeries.Count; s++)
            {
                Series series = chart.NSeries[s];
                ChartPointCollection points = series.Points;

                // Iterate over each point in the current series
                for (int p = 0; p < points.Count; p++)
                {
                    ChartPoint point = points[p];

                    // Assign a custom foreground color based on series and point index
                    if (s == 0) // First series
                    {
                        point.Area.ForegroundColor = (p % 2 == 0)
                            ? Color.FromArgb(79, 129, 189)
                            : Color.FromArgb(192, 80, 77);
                    }
                    else // Second series (or any additional series)
                    {
                        point.Area.ForegroundColor = (p % 2 == 0)
                            ? Color.FromArgb(100, 180, 100)
                            : Color.FromArgb(180, 100, 100);
                    }

                    // Ensure the formatting type is set to Custom so the color is applied
                    point.Area.Formatting = FormattingType.Custom;
                }
            }

            // Save the workbook with the customized chart
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ChartSeriesPointColorDemo.xlsx");
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                ChartSeriesPointColorDemo.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
