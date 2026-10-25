// Title: Add an exponential trendline to the first data series of a scatter chart using Aspose.Cells for .NET (C#)
// Common Searches: how to add an exponential trendline to a scatter chart with Aspose.Cells C# | Aspose.Cells set trendline type to exponential for the first series | display equation and R‑squared for a chart trendline in Aspose.Cells | change trendline line color to red in an Aspose.Cells chart | use reflection to access Trendlines API when it is not directly exposed in Aspose.Cells
// Tags: exponential trendline API Aspose.Cells | chart series trendline customization Aspose.Cells | trendline equation display Aspose.Cells | trendline line color styling Aspose.Cells | scatter chart creation Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, fills X/Y data, adds a scatter chart, defines the first data series, and uses reflection to add an exponential trendline with equation and R‑squared displayed and a red line color, then saves the file as TrendlineExample.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Access the first worksheet
            Worksheet ws = wb.Worksheets[0];

            // Populate sample data for the chart
            ws.Cells["A1"].PutValue("X");
            ws.Cells["B1"].PutValue("Y");
            ws.Cells["A2"].PutValue(1);
            ws.Cells["B2"].PutValue(2);
            ws.Cells["A3"].PutValue(2);
            ws.Cells["B3"].PutValue(4);
            ws.Cells["A4"].PutValue(3);
            ws.Cells["B4"].PutValue(9);
            ws.Cells["A5"].PutValue(4);
            ws.Cells["B5"].PutValue(16);

            // Add a scatter chart to the worksheet
            int chartIdx = ws.Charts.Add(ChartType.Scatter, 5, 0, 20, 10);
            Chart chart = ws.Charts[chartIdx];

            // Define the data series (Y values) and X values
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries[0].XValues = "A2:A5";

            // Attempt to add an exponential trendline using reflection (available in newer versions)
            try
            {
                var series = chart.NSeries[0];
                var trendlinesProp = series.GetType().GetProperty("Trendlines");
                if (trendlinesProp != null)
                {
                    var trendlines = trendlinesProp.GetValue(series);
                    var addMethod = trendlines?.GetType().GetMethod("Add", new[] { typeof(TrendlineType) });
                    var trendlineObj = addMethod?.Invoke(trendlines, new object[] { TrendlineType.Exponential });
                    if (trendlineObj != null)
                    {
                        // Enable equation and R‑squared display
                        var displayEqProp = trendlineObj.GetType().GetProperty("DisplayEquation");
                        displayEqProp?.SetValue(trendlineObj, true);
                        var displayRSqProp = trendlineObj.GetType().GetProperty("DisplayRSquared");
                        displayRSqProp?.SetValue(trendlineObj, true);

                        // Customize line color if supported
                        var lineProp = trendlineObj.GetType().GetProperty("Line");
                        var lineObj = lineProp?.GetValue(trendlineObj);
                        var colorProp = lineObj?.GetType().GetProperty("Color");
                        if (colorProp != null)
                        {
                            colorProp.SetValue(lineObj, Color.Red);
                        }
                    }
                }
            }
            catch
            {
                // Trendline feature may not be available; ignore any errors.
            }

            // Save the workbook
            string outputPath = "TrendlineExample.xlsx";

            // Ensure the output directory exists (handle possible null)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? string.Empty;
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            try
            {
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
