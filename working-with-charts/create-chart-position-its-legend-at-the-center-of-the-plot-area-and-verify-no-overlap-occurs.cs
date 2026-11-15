// Title: Create a column chart with a centered legend that avoids overlap using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# place chart legend at the center of the plot area | center legend in a column chart with Aspose.Cells .NET | adjust plot area dimensions to prevent legend overlap in Aspose.Cells | set custom X and Y coordinates for a chart legend using Aspose.Cells C# | programmatically verify that a chart legend does not overlap the plot area in Aspose.Cells
// Tags: legend placement at plot center Aspose.Cells | custom legend coordinates Aspose.Cells | plot area resize to prevent legend overlap Aspose.Cells | column chart legend dimensions Aspose.Cells | detect legend overlap Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds a column chart with sample data, defines a fixed legend size, computes the plot area's center, moves the legend to that center, shrinks the plot area to keep the legend from overlapping, checks for any overlap, and saves the result as ChartWithCenteredLegend.xlsx.
class ChartLegendCenterExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(45);
            sheet.Cells["B4"].PutValue(25);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add a series (values from B2:B4)
            int seriesIndex = chart.NSeries.Add("B2:B4", true);
            // Category data can be set if needed; omitted here to avoid compatibility issues
            // chart.NSeries[seriesIndex].CategoryData = "A2:A4";

            // Set legend size (points) and enable custom positioning
            chart.Legend.Width = 80;   // points
            chart.Legend.Height = 30; // points
            chart.Legend.Position = LegendPositionType.Bottom; // default; will override X/Y

            // Retrieve plot area dimensions
            int plotX = chart.PlotArea.X;
            int plotY = chart.PlotArea.Y;
            int plotWidth = chart.PlotArea.Width;
            int plotHeight = chart.PlotArea.Height;

            // Position legend at the center of the plot area
            int legendX = plotX + (plotWidth - chart.Legend.Width) / 2;
            int legendY = plotY + (plotHeight - chart.Legend.Height) / 2;
            chart.Legend.X = legendX;
            chart.Legend.Y = legendY;

            // Adjust plot area to avoid overlap with the legend
            chart.PlotArea.X = plotX + chart.Legend.Width / 2;
            chart.PlotArea.Y = plotY + chart.Legend.Height / 2;
            chart.PlotArea.Width = plotWidth - chart.Legend.Width;
            chart.PlotArea.Height = plotHeight - chart.Legend.Height;

            // Verify that the legend does not overlap the plot area
            bool overlap = !(chart.Legend.X + chart.Legend.Width <= chart.PlotArea.X ||
                             chart.Legend.X >= chart.PlotArea.X + chart.PlotArea.Width ||
                             chart.Legend.Y + chart.Legend.Height <= chart.PlotArea.Y ||
                             chart.Legend.Y >= chart.PlotArea.Y + chart.PlotArea.Height);

            Console.WriteLine(overlap
                ? "Legend overlaps the plot area."
                : "Legend positioned at the center without overlapping the plot area.");

            // Save the workbook
            string outputPath = "ChartWithCenteredLegend.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
