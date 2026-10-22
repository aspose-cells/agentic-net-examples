// Title: Add a linear trendline with equation and R‑squared to a line chart using Aspose.Cells for .NET (C#)
// AI Prompts: Create a line chart from worksheet data and programmatically attach a linear trendline that shows its equation and R‑squared value with Aspose.Cells in C#. | Employ reflection to retrieve the Trendlines collection of a chart series and set the ShowEquation and ShowRSquared properties for compatibility with older Aspose.Cells releases. | Save the workbook containing the configured chart to an XLSX file and confirm the file path.
// Common Searches: aspocells add linear trendline to line chart c# example | show equation and r squared on chart series using aspocells .net | how to use reflection to add trendline in older aspocells versions | c# aspocells line chart with regression equation | display r-squared value on aspocells chart
// Tags: Aspose.Cells add linear trendline to chart | show equation on Aspose.Cells chart series | display R-squared value in Aspose.Cells line chart | C# reflection access Trendlines collection | save workbook as XLSX with chart trendline

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new workbook, fills cells A1:B6 with quadratic data, adds a line chart, uses reflection to add a linear trendline to the series, enables ShowEquation and ShowRSquared, sets a chart title, and saves the workbook as LineChartWithTrendline.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the line chart (A1:B6)
            sheet.Cells["A1"].PutValue("X");
            sheet.Cells["B1"].PutValue("Y");
            for (int i = 2; i <= 6; i++)
            {
                sheet.Cells[$"A{i}"].PutValue(i - 1);               // X values: 1,2,3,4,5
                sheet.Cells[$"B{i}"].PutValue((i - 1) * (i - 1)); // Y values: 1,4,9,16,25 (quadratic)
            }

            // Add a line chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Line, 7, 0, 25, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart series
            chart.NSeries.Add("B2:B6", true);
            // Use XValues to specify category (X) data
            chart.NSeries[0].XValues = "A2:A6";

            // Attempt to add a linear trendline using reflection (compatible with older versions)
            try
            {
                var series = chart.NSeries[0];
                var trendlinesProp = series.GetType().GetProperty("Trendlines");
                if (trendlinesProp != null)
                {
                    var trendlines = trendlinesProp.GetValue(series);
                    var addMethod = trendlines.GetType().GetMethod("Add", new[] { typeof(TrendlineType) });
                    if (addMethod != null)
                    {
                        var trendline = addMethod.Invoke(trendlines, new object[] { TrendlineType.Linear });
                        var showEqProp = trendline.GetType().GetProperty("ShowEquation");
                        var showRSqProp = trendline.GetType().GetProperty("ShowRSquared");
                        showEqProp?.SetValue(trendline, true);
                        showRSqProp?.SetValue(trendline, true);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Trendline not added: {ex.Message}");
            }

            // Optional: give the chart a title
            chart.Title.Text = "Line Chart with Trendline";

            // Ensure output directory exists
            string outputPath = "LineChartWithTrendline.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
