// Title: How to assign distinct theme colors to each series in an Aspose.Cells column chart using C#
// AI Prompts: Generate C# code that loops through a Chart.NSeries collection in Aspose.Cells and sets a unique System.Drawing.Color for each series' Area.ForegroundColor and Border.Color, using a predefined color array. | Show a snippet that applies a cyclic color palette to chart series in Aspose.Cells, handling cases where the number of series exceeds the palette size.
// Common Searches: aspnet aspose.cells assign different colors to each chart series | c# set individual series fill color in Aspose.Cells column chart | loop through chart nseries to apply custom theme colors aspose.cells | apply repeating color palette to chart series when series count exceeds palette size aspose.cells | example of setting series border color in Aspose.Cells chart using c#
// Tags: set series area foregroundcolor Aspose.Cells | assign series border color Aspose.Cells | loop through chart series Aspose.Cells | apply cyclic color palette to chart series | column chart series color customization .NET

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, populates it with sample data, adds a column chart, defines an array of System.Drawing.Color values, iterates over the chart's NSeries collection, and assigns each series a distinct foreground and border color from the array before saving the workbook as SeriesThemeColors.xlsx.
class AssignSeriesThemeColors
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["C1"].PutValue("Series 2");
            sheet.Cells["D1"].PutValue("Series 3");
            sheet.Cells["E1"].PutValue("Series 4");
            sheet.Cells["F1"].PutValue("Series 5");

            for (int i = 2; i <= 6; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
                sheet.Cells[$"B{i}"].PutValue(i * 10);
                sheet.Cells[$"C{i}"].PutValue(i * 12);
                sheet.Cells[$"D{i}"].PutValue(i * 14);
                sheet.Cells[$"E{i}"].PutValue(i * 16);
                sheet.Cells[$"F{i}"].PutValue(i * 18);
            }

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 8, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add series to the chart (each column corresponds to a series)
            // The first column (A) is the category axis, so we start from B
            chart.NSeries.Add("B2:F6", true);

            // Define a set of theme colors to assign
            Color[] themeColors = new Color[]
            {
                Color.FromArgb(0x4F, 0x81, 0xBD), // Blue
                Color.FromArgb(0xC0, 0x50, 0x4D), // Red
                Color.FromArgb(0x9B, 0xBA, 0x59), // Green
                Color.FromArgb(0xFF, 0xC0, 0x00), // Orange
                Color.FromArgb(0x7F, 0x7F, 0x7F)  // Gray
            };

            // Assign a theme color to each series
            for (int i = 0; i < chart.NSeries.Count; i++)
            {
                // The correct type for a chart series in Aspose.Cells is Series
                Series series = chart.NSeries[i];

                // Set the fill color of the series area
                series.Area.ForegroundColor = themeColors[i % themeColors.Length];

                // Set the border color to match the fill color
                series.Border.Color = themeColors[i % themeColors.Length];
            }

            // Save the workbook
            workbook.Save("SeriesThemeColors.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
