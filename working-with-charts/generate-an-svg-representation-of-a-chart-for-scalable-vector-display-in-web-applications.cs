// Title: Export a column chart from an Aspose.Cells workbook to an SVG file with a transparent background using C#
// AI Prompts: Write C# code that builds a column chart from a worksheet range, sets a chart title and legend, and saves the chart as an SVG file with transparency using Aspose.Cells. | Adapt the example to create a line chart, move the legend to the bottom, specify custom image dimensions, and render the chart to SVG in .NET. | Add data labels to each series, adjust the chart size, and export the updated chart to an SVG file with Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# render Excel chart as SVG with transparent background | How to export a column chart to SVG using Aspose.Cells ImageOrPrintOptions | Save Aspose.Cells chart as scalable vector graphic instead of PNG in .NET | C# example for converting worksheet chart to SVG with Aspose.Cells | Aspose.Cells chart ToImage SVG output tutorial
// Tags: Aspose.Cells chart to SVG conversion | C# column chart rendering with Aspose.Cells | ImageOrPrintOptions SVG transparency | Aspose.Cells ToImage method for chart export | Aspose.Cells chart legend placement

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsSvgChartExample
{
    // The sample creates a workbook, fills cells with category and series data, adds a column chart referencing those ranges, configures the title, series names, category axis, and legend position, then uses ImageOrPrintOptions to render the chart directly to a transparent‑background SVG file named 'ChartOutput.svg' via Aspose.Cells' ToImage method.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                // A1:A5 – Categories
                // B1:B5 – Series 1 values
                // C1:C5 – Series 2 values
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Series 1");
                sheet.Cells["C1"].PutValue("Series 2");

                string[] categories = { "Jan", "Feb", "Mar", "Apr", "May" };
                double[] series1 = { 10, 20, 30, 25, 15 };
                double[] series2 = { 15, 25, 20, 30, 10 };

                for (int i = 0; i < categories.Length; i++)
                {
                    sheet.Cells[i + 1, 0].PutValue(categories[i]);   // Column A
                    sheet.Cells[i + 1, 1].PutValue(series1[i]);    // Column B
                    sheet.Cells[i + 1, 2].PutValue(series2[i]);    // Column C
                }

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set chart title
                chart.Title.Text = "Sample Column Chart";

                // Add first series (Series 1)
                int series1Index = chart.NSeries.Add("B2:B6", true);
                chart.NSeries[series1Index].Name = "Series 1";

                // Add second series (Series 2)
                int series2Index = chart.NSeries.Add("C2:C6", true);
                chart.NSeries[series2Index].Name = "Series 2";

                // Set category axis data (X‑axis)
                chart.NSeries.CategoryData = "A2:A6";

                // Optional: format the chart (e.g., legend position)
                chart.Legend.Position = LegendPositionType.Right;

                // Prepare image options for SVG output
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    // Define a transparent background (optional)
                    Transparent = true
                };

                // Render the chart directly to an SVG file
                chart.ToImage("ChartOutput.svg", imgOptions);

                // Optionally save the entire workbook
                // workbook.Save("WorkbookWithChart.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
