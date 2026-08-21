// Title: Create a Colored Pyramid Chart with Aspose.Cells for .NET
// Description: Demonstrates how to build a new Workbook, populate category and series data, add a Pyramid chart, assign a unique solid fill color to each series (level) via the series Area ForegroundColor, and save the file as PyramidChartWithDistinctColors.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells pyramid chart | C# pyramid chart example | custom series colors Aspose.Cells | Excel chart level colors | set ForegroundColor Aspose.Cells | .NET chart customization | colored pyramid chart
// Common Searches: Aspose.Cells how to color pyramid chart levels | C# example for pyramid chart with custom colors | set series fill color Aspose.Cells chart | create pyramid chart from range B2:D5 Aspose.Cells | assign different colors to each pyramid layer
// Developer Intent: Generate a pyramid chart and apply a distinct color to each level using Aspose.Cells for .NET.
// Use Cases: Show a sales funnel where each tier is highlighted with brand‑specific colors. | Produce presentation‑ready Excel reports that require clear visual separation of pyramid layers. | Automate recurring dashboards that render hierarchical data with easily distinguishable colors.
// AI Prompts: Write C# code with Aspose.Cells to create a pyramid chart and set a gradient fill for each level instead of solid colors. | Explain how to add data labels, a legend, and tooltips to the colored pyramid chart while keeping the custom series colors. | Provide a snippet that reads data from a database and dynamically assigns colors to pyramid levels based on a configuration file.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace PyramidChartDemo
{
    // Demonstrates how to build a new Workbook, populate category and series data, add a Pyramid chart, assign a unique solid fill color to each series (level) via the series Area ForegroundColor, and save the file as PyramidChartWithDistinctColors.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                // Categories
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Level 1");
                sheet.Cells["A3"].PutValue("Level 2");
                sheet.Cells["A4"].PutValue("Level 3");
                sheet.Cells["A5"].PutValue("Level 4");

                // Series values (each series will become a separate level in the pyramid)
                sheet.Cells["B1"].PutValue("Series 1");
                sheet.Cells["C1"].PutValue("Series 2");
                sheet.Cells["D1"].PutValue("Series 3");

                sheet.Cells["B2"].PutValue(40);
                sheet.Cells["B3"].PutValue(30);
                sheet.Cells["B4"].PutValue(20);
                sheet.Cells["B5"].PutValue(10);

                sheet.Cells["C2"].PutValue(35);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["C4"].PutValue(15);
                sheet.Cells["C5"].PutValue(5);

                sheet.Cells["D2"].PutValue(30);
                sheet.Cells["D3"].PutValue(20);
                sheet.Cells["D4"].PutValue(10);
                sheet.Cells["D5"].PutValue(5);

                // Add a Pyramid chart
                int chartIndex = sheet.Charts.Add(ChartType.Pyramid, 6, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart (all series)
                chart.NSeries.Add("B2:D5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Assign distinct colors to each series (level)
                Color[] levelColors = new Color[] { Color.Red, Color.Green, Color.Blue };
                for (int i = 0; i < chart.NSeries.Count && i < levelColors.Length; i++)
                {
                    Series series = chart.NSeries[i];
                    // Use ForegroundColor to apply a solid fill color to the series
                    series.Area.ForegroundColor = levelColors[i];
                }

                // Save the workbook
                workbook.Save("PyramidChartWithDistinctColors.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
