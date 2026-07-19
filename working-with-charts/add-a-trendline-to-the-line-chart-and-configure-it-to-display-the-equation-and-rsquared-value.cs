// Title: Add a Linear Trendline with Equation & R‑Squared to a Line Chart – Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills sample X/Y data, inserts a line chart, adds a linear trendline to the first series, and configures the trendline to show its regression equation and R‑squared value. The example also demonstrates setting a custom name and blue color before saving the file as an XLSX workbook.
// Keywords: Aspose.Cells | C# | .NET charting | line chart trendline | linear regression trendline | display equation | show R‑squared | trendline color | trendline name | chart customization
// Common Searches: Aspose.Cells add linear trendline to chart | show equation and R‑squared on Aspose.Cells chart | customize trendline color in Aspose.Cells C# | set trendline name Aspose.Cells line chart | C# example for trendline in Aspose.Cells
// Developer Intent: Add a linear trendline to a line chart and enable its equation and R‑squared display.
// Use Cases: Generate a line chart from worksheet data and attach a regression line for quick data analysis. | Present the regression formula and goodness‑of‑fit (R²) directly on the chart for reports or dashboards. | Match corporate styling by assigning a specific name and color to the trendline.
// AI Prompts: Write C# code using Aspose.Cells to add a polynomial trendline with its equation and R‑squared to a chart. | Show how to add multiple trendlines of different types to separate series in the same Aspose.Cells chart. | Explain how to format the trendline equation text (font, size, color) in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsTrendlineExample
{
    // Creates a workbook, fills sample X/Y data, inserts a line chart, adds a linear trendline to the first series, and configures the trendline to show its regression equation and R‑squared value. The example also demonstrates setting a custom name and blue color before saving the file as an XLSX workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the line chart
            worksheet.Cells["A1"].PutValue(1);
            worksheet.Cells["A2"].PutValue(2);
            worksheet.Cells["A3"].PutValue(3);
            worksheet.Cells["A4"].PutValue(4);
            worksheet.Cells["B1"].PutValue(2);
            worksheet.Cells["B2"].PutValue(4);
            worksheet.Cells["B3"].PutValue(6);
            worksheet.Cells["B4"].PutValue(8);

            // Add a line chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Line, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the series data (Y values) and category (X values)
            chart.NSeries.Add("B1:B4", true);
            chart.NSeries.CategoryData = "A1:A4";

            // Add a linear trendline to the first series
            int trendlineIdx = chart.NSeries[0].TrendLines.Add(TrendlineType.Linear);
            Trendline trendline = chart.NSeries[0].TrendLines[trendlineIdx];

            // Configure the trendline to display the equation and R‑squared value
            trendline.DisplayEquation = true;
            trendline.DisplayRSquared = true;

            // Optional: set a custom name or color for the trendline
            trendline.Name = "Linear Trend";
            trendline.Color = Color.Blue;

            // Save the workbook to a file
            workbook.Save("LineChartWithTrendline.xlsx");
        }
    }
}
