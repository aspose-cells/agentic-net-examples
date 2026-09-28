// Title: How to add a linear trendline that shows the equation and R‑squared value to a line chart using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a line chart with Aspose.Cells, adds a linear trendline, and enables the equation and R‑squared display. | Show how to configure a trendline's name and color while displaying its regression equation and R² value in an Aspose.Cells chart.
// Common Searches: aspocells add linear trendline with equation and r2 to line chart c# | c# aspocells display regression equation on chart trendline | how to show r‑squared value for trendline in aspocells line chart | customize trendline name and color aspocells c# example | aspocells line chart linear regression line with statistics
// Tags: Aspose.Cells linear trendline | Aspose.Cells trendline equation display | Aspose.Cells trendline R-squared | Aspose.Cells trendline color customization | Aspose.Cells line chart regression

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsTrendlineExample
{
    // Demonstrates creating a workbook, inserting sample data, adding a line chart, attaching a linear trendline, enabling equation and R‑squared display, customizing the trendline's name and color, and saving the result as LineChartWithTrendline.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the line chart
            sheet.Cells["A1"].PutValue(1);
            sheet.Cells["A2"].PutValue(2);
            sheet.Cells["A3"].PutValue(3);
            sheet.Cells["A4"].PutValue(4);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["B2"].PutValue(4);
            sheet.Cells["B3"].PutValue(6);
            sheet.Cells["B4"].PutValue(8);

            // Add a line chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

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
