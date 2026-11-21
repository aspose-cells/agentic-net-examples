// Title: How to assign individual theme colors to each series in an Aspose.Cells column chart (C#)
// Common Searches: C# Aspose.Cells set different fill colors for each column chart series | apply custom theme colors to chart series using Aspose.Cells API | iterate NSeries collection to change series colors in an Aspose.Cells workbook | Aspose.Cells column chart series color customization example in .NET | programmatically change series foreground color in an Aspose.Cells chart
// Tags: Aspose.Cells set series foreground color | column chart series color customization Aspose.Cells | C# iterate NSeries collection Aspose.Cells | apply theme colors to chart series .NET | custom series colors Excel workbook Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, fills it with sample data, adds a column chart, defines an array of distinct colors, loops through the chart's NSeries collection, assigns each series' area foreground color from the array (wrapping when needed), and saves the file as SeriesThemeColors.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Populate sample data
            var cells = workbook.Worksheets[0].Cells;
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Series1");
            cells["C1"].PutValue("Series2");
            cells["D1"].PutValue("Series3");
            cells["A2"].PutValue("A");
            cells["A3"].PutValue("B");
            cells["A4"].PutValue("C");
            cells["B2"].PutValue(10);
            cells["B3"].PutValue(20);
            cells["B4"].PutValue(30);
            cells["C2"].PutValue(15);
            cells["C3"].PutValue(25);
            cells["C4"].PutValue(35);
            cells["D2"].PutValue(12);
            cells["D3"].PutValue(22);
            cells["D4"].PutValue(32);

            // Add a column chart
            int chartIdx = workbook.Worksheets[0].Charts.Add(ChartType.Column, 5, 0, 20, 10);
            var chart = workbook.Worksheets[0].Charts[chartIdx];

            // Add series (one for each data column)
            chart.NSeries.Add("B2:D4", true);

            // Define a set of distinct colors to apply to each series
            Color[] seriesColors = new Color[]
            {
                Color.FromArgb(0x5B, 0x9B, 0xD5), // Accent1
                Color.FromArgb(0xED, 0x7D, 0x31), // Accent2
                Color.FromArgb(0xA5, 0xA5, 0xA5), // Accent3
                Color.FromArgb(0x70, 0xAD, 0x47), // Accent4
                Color.FromArgb(0xFF, 0xC0, 0x00), // Accent5
                Color.FromArgb(0x5C, 0x5C, 0x5C)  // Accent6
            };

            // Apply colors to each series
            for (int i = 0; i < chart.NSeries.Count; i++)
            {
                var series = chart.NSeries[i];
                Color fillColor = seriesColors[i % seriesColors.Length];
                series.Area.ForegroundColor = fillColor;
            }

            // Save the workbook
            string outputPath = "SeriesThemeColors.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
