// Title: How to set custom foreground colors for each chart point in an Aspose.Cells column chart using C#
// Common Searches: Aspose.Cells C# set individual column chart point color | change fill color of specific chart points using Aspose.Cells API | apply alternating foreground colors to chart points in .NET Excel chart | set solid fill pattern and border for chart points with Aspose.Cells | programmatically customize chart point appearance in Excel with C#
// Tags: chart point fill color Aspose.Cells C# | set chart point foreground color .NET | solid fill pattern chart points Aspose.Cells | chart point border styling Aspose.Cells | column chart series point customization Excel C#

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace ChartPointForegroundColorDemo
{
    // The example creates a workbook, adds sample data, inserts a column chart, and then iterates through each series and its points to assign alternating red and green foreground colors using a solid fill pattern, while also applying a thin black border to every chart point before saving the file as an XLSX workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
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

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["B5"].PutValue(40);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the series
                chart.NSeries.Add("B2:B5", true);          // Values
                chart.NSeries.CategoryData = "A2:A5";      // Categories
                chart.NSeries[0].Name = "Sample Series";

                // Iterate over each series and customize its points
                for (int s = 0; s < chart.NSeries.Count; s++)
                {
                    Series series = chart.NSeries[s];
                    ChartPointCollection points = series.Points;

                    for (int p = 0; p < points.Count; p++)
                    {
                        ChartPoint point = points[p];

                        // Alternate colors: even index -> Red, odd index -> Green
                        Color customColor = (p % 2 == 0) ? Color.Red : Color.Green;

                        // Use solid fill pattern
                        point.Area.FillFormat.Pattern = FillPattern.Solid;

                        // Set the foreground color via PatternFill
                        point.Area.FillFormat.PatternFill.ForegroundColor = customColor;

                        // Add a black border with thin weight
                        point.Border.Color = Color.Black;
                        point.Border.Weight = WeightType.SingleLine;
                    }
                }

                // Save the workbook
                string outputPath = "ChartPointForegroundColorDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
