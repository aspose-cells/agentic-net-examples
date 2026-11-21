// Title: Create a line chart with a linear trendline that displays the regression equation and R‑squared value using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add linear trendline to line chart and display equation | How to show R squared on an Excel line chart with Aspose.Cells .NET | C# code to add regression line with equation to chart using Aspose.Cells | Generate line chart with trendline and statistics using Aspose.Cells for .NET | Add trendline with equation and R2 to chart in Aspose.Cells C#
// Tags: add linear trendline Aspose.Cells C# | display regression equation Aspose.Cells chart | show R-squared Aspose.Cells line chart | configure chart trendline Aspose.Cells .NET | export chart with trendline to XLSX using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, fills cells with sample X/Y data, adds a line chart, inserts a linear trendline for the first series, and enables the display of both the regression equation and the R‑squared value before saving the file as an XLSX workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the line chart
            sheet.Cells["A1"].PutValue("X");
            sheet.Cells["B1"].PutValue("Y");
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(i + 1);                     // X values
                sheet.Cells[i + 1, 1].PutValue((i + 1) * (i + 2) / 2);   // Y values (example)
            }

            // Add a line chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 25, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B11", true);          // Y values
            // Use dynamic to avoid compile‑time binding issues with older API versions
            dynamic series = chart.NSeries[0];
            series.CategoryData = "A2:A11";            // X values

            // Add a linear trendline to the first series
            var trendline = series.Trendlines.Add(TrendlineType.Linear);
            trendline.DisplayEquation = true;
            trendline.DisplayRSquared = true;

            // Save the workbook with the chart and trendline
            string outputFile = "LineChartWithTrendline.xlsx";
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputFile)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
