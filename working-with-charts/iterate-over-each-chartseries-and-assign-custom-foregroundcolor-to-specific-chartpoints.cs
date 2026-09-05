// Title: How to set custom foreground colors for individual chart points in an Aspose.Cells column chart using C#
// AI Prompts: Write C# code with Aspose.Cells that loops through each series in a column chart and assigns a distinct ForegroundColor to every ChartPoint. | Show how to change the area color of chart points based on their index (e.g., first point red, second green, others blue) and then save the workbook. | Provide a complete example that creates sample data, adds a column chart, and customizes the foreground colors of individual data points programmatically.
// Common Searches: Aspose.Cells C# set foreground color for individual chart points in a column chart | how to change color of specific data points in an Excel chart using Aspose.Cells | loop through chart series and points to apply custom colors with Aspose.Cells .NET | example of customizing chart point area colors in Aspose.Cells C# | assign different colors to each data point in a column chart programmatically Aspose.Cells
// Tags: set chart point foreground color Aspose.Cells C# | iterate chart series points Aspose.Cells | customize column chart data point colors | apply individual point colors Excel chart Aspose | foreground color for chart point area Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

// The example creates a workbook, adds sample data, inserts a column chart, then iterates through each series and its points, assigning red, green, or blue foreground colors to the point areas based on the point index, and finally saves the file as SeriesPointsForegroundColor.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a column chart
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the series and categories
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Iterate over each series in the chart
        for (int seriesIdx = 0; seriesIdx < chart.NSeries.Count; seriesIdx++)
        {
            Series series = chart.NSeries[seriesIdx];

            // Iterate over each point in the current series
            for (int pointIdx = 0; pointIdx < series.Points.Count; pointIdx++)
            {
                ChartPoint point = series.Points[pointIdx];

                // Assign custom foreground colors based on point index
                if (pointIdx == 0)
                {
                    point.Area.ForegroundColor = Color.Red;
                }
                else if (pointIdx == 1)
                {
                    point.Area.ForegroundColor = Color.Green;
                }
                else
                {
                    point.Area.ForegroundColor = Color.Blue;
                }
            }
        }

        // Save the workbook with the customized chart
        workbook.Save("SeriesPointsForegroundColor.xlsx");
    }
}
