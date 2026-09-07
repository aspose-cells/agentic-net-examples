// Title: How to apply threshold‑based conditional colors to a progress bar chart in Excel using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a bar chart and assigns red, yellow, or green fill colors to each bar based on its numeric progress value. | Show how to read values from worksheet cells and conditionally set the foreground color of individual chart points using the Aspose.Cells .NET API. | Generate a complete example that builds an Excel file, adds a progress‑bar style bar chart, and applies threshold‑driven color formatting to the chart series.
// Common Searches: Aspose.Cells C# set bar chart point color based on cell value | conditional formatting of chart series in Aspose.Cells .NET example | change individual bar colors in Excel chart using Aspose.Cells API | progress bar visualization with red yellow green colors in Aspose.Cells
// Tags: Aspose.Cells conditional chart point coloring | C# bar chart color thresholds Aspose.Cells | Excel progress bar chart formatting Aspose.Cells | set chart series point foreground color .NET | threshold based bar colors Excel C#

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ProgressBarExample
{
    // Demonstrates creating an Excel workbook with Aspose.Cells, adding a bar chart that acts as a progress bar, and applying conditional fill colors (red, yellow, green) to each bar based on progress values using C#.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the progress bar
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["B1"].PutValue("Progress");
                sheet.Cells["A2"].PutValue("Task 1");
                sheet.Cells["B2"].PutValue(20);   // 20% progress
                sheet.Cells["A3"].PutValue("Task 2");
                sheet.Cells["B3"].PutValue(55);   // 55% progress
                sheet.Cells["A4"].PutValue("Task 3");
                sheet.Cells["B4"].PutValue(85);   // 85% progress

                // Add a bar chart that will act as a progress bar
                int chartIndex = sheet.Charts.Add(ChartType.Bar, 6, 0, 20, 7);
                Chart chart = sheet.Charts[chartIndex];
                chart.Title.Text = "Progress Bar";

                // Define the series (values) and categories (tasks)
                int dataCount = 3; // number of tasks
                string valuesRange = $"B2:B{dataCount + 1}";
                string categoriesRange = $"A2:A{dataCount + 1}";

                chart.NSeries.Add(valuesRange, true);
                chart.NSeries.CategoryData = categoriesRange;

                // Threshold values for conditional coloring
                double redThreshold = 30;    // below 30% -> red
                double yellowThreshold = 70; // 30% - 70% -> yellow
                // above 70% -> green

                // Apply conditional colors to each data point based on its value
                var series = chart.NSeries[0]; // Series object
                for (int i = 0; i < series.Points.Count; i++)
                {
                    // Retrieve the progress value from the worksheet (row i+2, column B)
                    double progressValue = sheet.Cells[i + 1, 1].DoubleValue;

                    // Determine fill color according to thresholds
                    Color fillColor;
                    if (progressValue < redThreshold)
                        fillColor = Color.Red;
                    else if (progressValue <= yellowThreshold)
                        fillColor = Color.Yellow;
                    else
                        fillColor = Color.Green;

                    // Set the fill color of the corresponding bar (data point)
                    series.Points[i].Area.ForegroundColor = fillColor;
                }

                // Save the workbook with the conditional formatted progress bar chart
                string outputPath = "ProgressBarConditionalFormatting.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
