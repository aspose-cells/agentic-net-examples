// Title: How to create a reusable Excel chart template with Aspose.Cells for .NET and apply its styling to new charts
// Common Searches: Aspose.Cells C# create chart template workbook and reuse it for other charts | Copy chart formatting from a saved Excel template using Aspose.Cells .NET | Apply source chart style to a new column chart in a different workbook with Aspose.Cells | Reuse series formatting from a template chart in Aspose.Cells C# example | Load chart template file and generate styled charts programmatically with Aspose.Cells
// Tags: aspocells chart template generation | aspocells clone chart attributes | aspocells transfer chart styling programmatically | aspocells copy series visual settings | aspocells create styled Excel charts via code

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example demonstrates how to build a workbook that contains a formatted column chart, save it as a template, then load that template in another workbook, copy the chart's type, titles, fonts, axis titles, and series visual settings, reassign the data source, and save the new workbook with the transferred styling.
class ChartTemplateExample
{
    static void Main()
    {
        try
        {
            // -----------------------------
            // Step 1: Create a workbook with a sample chart and save it as a template workbook
            // -----------------------------
            Workbook templateWb = new Workbook();
            Worksheet templateWs = templateWb.Worksheets[0];

            // Populate sample data
            templateWs.Cells["A1"].PutValue("Category");
            templateWs.Cells["B1"].PutValue("Value");
            templateWs.Cells["A2"].PutValue("Jan");
            templateWs.Cells["A3"].PutValue("Feb");
            templateWs.Cells["A4"].PutValue("Mar");
            templateWs.Cells["B2"].PutValue(120);
            templateWs.Cells["B3"].PutValue(150);
            templateWs.Cells["B4"].PutValue(180);

            // Add a chart
            int chartIdx = templateWs.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            Chart templateChart = templateWs.Charts[chartIdx];

            // Set titles and formatting
            templateChart.Title.Text = "Sales Overview";
            templateChart.Title.Font.Color = Color.DarkBlue;
            templateChart.Title.Font.Size = 14;
            templateChart.CategoryAxis.Title.Text = "Month";
            templateChart.ValueAxis.Title.Text = "Revenue";

            // Add series
            int seriesIdx = templateChart.NSeries.Add("B2:B4", true);
            templateChart.NSeries[seriesIdx].Name = "Revenue";
            templateChart.NSeries[seriesIdx].Area.ForegroundColor = Color.CornflowerBlue;
            templateChart.NSeries[seriesIdx].Area.Transparency = 0.2;

            // Save the workbook that contains the chart (used later as a template source)
            string templateFile = "TemplateWorkbook.xlsx";
            templateWb.Save(templateFile);

            // -----------------------------
            // Step 2: Create a new workbook and apply the saved chart template to a new chart
            // -----------------------------
            Workbook newWb = new Workbook();
            Worksheet newWs = newWb.Worksheets[0];

            // Populate data for the new chart
            newWs.Cells["A1"].PutValue("Category");
            newWs.Cells["B1"].PutValue("Value");
            newWs.Cells["A2"].PutValue("Apr");
            newWs.Cells["A3"].PutValue("May");
            newWs.Cells["A4"].PutValue("Jun");
            newWs.Cells["B2"].PutValue(200);
            newWs.Cells["B3"].PutValue(170);
            newWs.Cells["B4"].PutValue(210);

            // Load the template workbook if it exists
            if (!File.Exists(templateFile))
                throw new FileNotFoundException($"Template file '{templateFile}' not found.");

            Workbook templateSourceWb = new Workbook(templateFile);
            Worksheet templateSourceWs = templateSourceWb.Worksheets[0];

            if (templateSourceWs.Charts.Count == 0)
                throw new InvalidOperationException("No chart found in the template workbook.");

            // Get the first chart from the template workbook
            Chart sourceChart = templateSourceWs.Charts[0];

            // Create a new chart in the destination worksheet
            int newChartIdx = newWs.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            Chart newChart = newWs.Charts[newChartIdx];

            // Manually copy required properties from the source chart
            newChart.Type = sourceChart.Type;
            newChart.Title.Text = sourceChart.Title.Text;
            newChart.Title.Font.Color = sourceChart.Title.Font.Color;
            newChart.Title.Font.Size = sourceChart.Title.Font.Size;
            newChart.CategoryAxis.Title.Text = sourceChart.CategoryAxis.Title.Text;
            newChart.ValueAxis.Title.Text = sourceChart.ValueAxis.Title.Text;

            // Update the series data source to point to the new worksheet data
            newChart.NSeries.Clear();
            int newSeriesIdx = newChart.NSeries.Add("B2:B4", true);
            newChart.NSeries[newSeriesIdx].Name = "Revenue Q2";

            // Copy series formatting from the source chart (if any)
            if (sourceChart.NSeries.Count > 0)
            {
                var srcSeries = sourceChart.NSeries[0];
                var destSeries = newChart.NSeries[newSeriesIdx];
                destSeries.Area.ForegroundColor = srcSeries.Area.ForegroundColor;
                destSeries.Area.Transparency = srcSeries.Area.Transparency;
            }

            // Save the workbook with the styled chart
            newWb.Save("StyledChartWorkbook.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
