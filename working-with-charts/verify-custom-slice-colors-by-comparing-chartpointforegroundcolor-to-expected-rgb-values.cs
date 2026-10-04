// Title: Verify Excel chart slice colors with Aspose.Cells for .NET by comparing ChartPoint.Area.ForegroundColor to expected RGB values
// AI Prompts: Write C# code using Aspose.Cells that opens a workbook, locates the first chart, iterates through each series point, reads point.Area.ForegroundColor, and checks it against a dictionary of expected Color values, printing match or mismatch results. | Update the sample to ignore points that have no predefined expected color, log a warning for those slices, and continue reporting matches for the defined colors.
// Common Searches: aspocells c# read pie chart slice color | how to compare chart point foreground color with expected RGB in Aspose.Cells | validate custom colors of Excel chart slices using Aspose.Cells .NET | retrieve ChartPoint.Area.ForegroundColor in C# Aspose.Cells | check if Excel chart slice colors match predefined palette Aspose
// Tags: Aspose.Cells chart slice color verification | C# ChartPoint foreground color comparison | Excel pie chart color validation Aspose | read chart point area color .NET | compare chart point RGB Aspose.Cells

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an Excel workbook, accesses the first chart on the first worksheet, defines expected RGB colors for each slice index, iterates through all series points, reads each point's Area.ForegroundColor, compares it to the expected color, and writes match or mismatch information to the console.
class VerifyChartSliceColors
{
    static void Main()
    {
        try
        {
            const string inputFile = "InputChart.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: File \"{inputFile}\" not found.");
                return;
            }

            // Load the workbook containing the chart
            Workbook workbook = new Workbook(inputFile);

            // Assume the chart is on the first worksheet and is the first chart object
            Worksheet sheet = workbook.Worksheets[0];
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the first worksheet.");
                return;
            }

            Chart chart = sheet.Charts[0];

            // Expected colors for each slice (point index -> expected RGB color)
            Dictionary<int, Color> expectedColors = new Dictionary<int, Color>()
            {
                { 0, Color.FromArgb(255, 255, 0, 0) },   // Red
                { 1, Color.FromArgb(255, 0, 255, 0) },   // Green
                { 2, Color.FromArgb(255, 0, 0, 255) }    // Blue
                // Add more expected colors as needed
            };

            // Iterate through each series in the chart
            foreach (Series series in chart.NSeries)
            {
                // Iterate through each point (slice) in the series
                for (int i = 0; i < series.Points.Count; i++)
                {
                    ChartPoint point = series.Points[i];

                    // Retrieve the actual foreground color of the slice via the point's Area
                    Color actualColor = point.Area.ForegroundColor;

                    // Determine the expected color for this point index
                    if (expectedColors.TryGetValue(i, out Color expectedColor))
                    {
                        bool colorsMatch = actualColor.ToArgb() == expectedColor.ToArgb();

                        Console.WriteLine($"Slice {i}: " +
                                          $"Actual = #{actualColor.ToArgb():X8}, " +
                                          $"Expected = #{expectedColor.ToArgb():X8} => " +
                                          (colorsMatch ? "Match" : "Mismatch"));
                    }
                    else
                    {
                        Console.WriteLine($"Slice {i}: No expected color defined for this index.");
                    }
                }
            }

            // No need to save the workbook as we are only verifying colors
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
