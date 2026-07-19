// Title: Add Exponential Trendline to First Series of a Line Chart with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills two columns with sample data, builds a line chart, binds X‑ and Y‑values, inserts an exponential trendline on the first series, shows the regression equation and R², colors the line red, and saves the workbook.
// Keywords: Aspose.Cells C# trendline | exponential trendline Aspose.Cells | chart series trendline .NET | display regression equation Aspose.Cells | set trendline color Aspose.Cells | line chart trendline example | Aspose.Cells GitHub sample
// Common Searches: Aspose.Cells add exponential trendline C# | how to show trendline equation and R squared in Aspose.Cells chart | set trendline color in Aspose.Cells line chart | C# Aspose.Cells chart series trendline types | example of trendline in Aspose.Cells workbook
// Developer Intent: Insert an exponential trendline into the first data series of a line chart and configure its visual and analytical properties.
// Use Cases: Model growth patterns by applying an exponential regression line to chart data. | Present regression formula and R² directly on the chart for quick insight. | Highlight the trendline with a custom color to improve visual distinction.
// AI Prompts: Generate C# code using Aspose.Cells to add a polynomial trendline to the second series of a column chart and hide the equation label. | Create an example that adds a logarithmic trendline to a scatter chart, applies a dashed line style, and exports the workbook to PDF. | Provide step‑by‑step instructions for adding multiple trendlines of different types (linear, exponential) to separate series in the same chart with Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace TrendlineExample
{
    // Creates a workbook, fills two columns with sample data, builds a line chart, binds X‑ and Y‑values, inserts an exponential trendline on the first series, shows the regression equation and R², colors the line red, and saves the workbook.
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

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue(1);
                sheet.Cells["A2"].PutValue(2);
                sheet.Cells["A3"].PutValue(3);
                sheet.Cells["A4"].PutValue(4);
                sheet.Cells["B1"].PutValue(2);
                sheet.Cells["B2"].PutValue(4);
                sheet.Cells["B3"].PutValue(6);
                sheet.Cells["B4"].PutValue(8);

                // Add a line chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the first series (Y values)
                chart.NSeries.Add("B1:B4", true);

                // Set the category (X) data using XValues property
                chart.NSeries[0].XValues = "A1:A4";

                // Insert an exponential trendline into the first data series
                int trendlineIdx = chart.NSeries[0].TrendLines.Add(TrendlineType.Exponential);
                Trendline trendline = chart.NSeries[0].TrendLines[trendlineIdx];

                // Display equation, R‑squared value, and set line color
                trendline.DisplayEquation = true;
                trendline.DisplayRSquared = true;
                trendline.Color = Color.Red;

                // Save the workbook to a file
                workbook.Save("TrendlineExponential.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
