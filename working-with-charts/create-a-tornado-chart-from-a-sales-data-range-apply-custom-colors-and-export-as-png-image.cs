// Title: Create a Tornado (Stacked Bar) Chart with Custom Colors and Export to PNG using Aspose.Cells for .NET (C#)
// Description: This C# example shows how to build a workbook, add regional sales data, generate a stacked‑bar tornado chart, apply red and blue colors to the series, and export the chart as a PNG image with Aspose.Cells for .NET.
// Keywords: Aspose.Cells tornado chart C# | stacked bar chart custom colors .NET | export chart to PNG Aspose.Cells | sales data tornado diagram | ChartType.BarStacked Aspose.Cells | chart.ToImage example | C# Excel chart generation | regional sales comparison chart
// Common Searches: how to create a tornado chart with Aspose.Cells | Aspose.Cells stacked bar chart custom colors | export Aspose.Cells chart as PNG in C# | sample code for tornado chart using Aspose.Cells | set chart data range Aspose.Cells .NET
// Developer Intent: Generate a tornado‑style stacked bar chart from sales figures, color each series uniquely, and save the chart as a PNG image using Aspose.Cells for .NET.
// Use Cases: Compare year‑over‑year sales across regions in a presentation‑ready tornado chart. | Produce colored comparative charts for quarterly performance reports that can be embedded in PDFs or PowerPoint. | Automate creation of PNG chart images for a web dashboard that visualizes regional sales trends.
// AI Prompts: Add data labels to each bar of the tornado chart using Aspose.Cells in C#. | Show how to reverse the category axis order for a tornado chart with Aspose.Cells. | Provide code to export the chart to JPEG or SVG instead of PNG using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This C# example shows how to build a workbook, add regional sales data, generate a stacked‑bar tornado chart, apply red and blue colors to the series, and export the chart as a PNG image with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add header row
            sheet.Cells["A1"].PutValue("Region");
            sheet.Cells["B1"].PutValue("Sales 2022");
            sheet.Cells["C1"].PutValue("Sales 2023");

            // Sample data for a tornado chart (one series positive, one negative)
            string[] regions = { "North", "South", "East", "West" };
            double[] sales2022 = { 120, 150, 100, 130 };   // Positive values
            double[] sales2023 = { -80, -110, -70, -90 };  // Negative values for opposite side

            for (int i = 0; i < regions.Length; i++)
            {
                sheet.Cells[i + 2, 0].PutValue(regions[i]);          // Region name
                sheet.Cells[i + 2, 1].PutValue(sales2022[i]);       // 2022 sales
                sheet.Cells[i + 2, 2].PutValue(sales2023[i]);       // 2023 sales (negative)
            }

            // Add a stacked bar chart (tornado style)
            int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Define the data range for the chart (including categories)
            chart.SetChartDataRange("A1:C5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Apply custom colors to each series
            chart.NSeries[0].Area.ForegroundColor = Color.Red;   // 2022 series
            chart.NSeries[1].Area.ForegroundColor = Color.Blue;  // 2023 series (negative side)

            // (Optional) Reverse category order – not supported in this API version
            // chart.CategoryAxis.IsInversed = true;

            // Export the chart as a PNG image
            chart.ToImage("tornado_chart.png");

            // Save the workbook (optional, for verification)
            workbook.Save("tornado_chart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
