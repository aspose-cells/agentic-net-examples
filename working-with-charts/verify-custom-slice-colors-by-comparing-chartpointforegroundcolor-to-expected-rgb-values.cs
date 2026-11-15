// Title: Validate custom pie chart slice colors in an Excel workbook with Aspose.Cells for .NET by comparing ChartPoint.Area.ForegroundColor to expected RGB values
// Common Searches: Aspose.Cells C# read pie chart slice foreground color | compare Excel chart slice color with expected RGB using Aspose.Cells | verify pie chart slice colors programmatically in .NET | how to get ChartPoint foreground color in Aspose.Cells | check if Excel chart slice colors match predefined colors in C#
// Tags: Aspose.Cells chart slice color verification | ChartPoint Area ForegroundColor retrieval | compare chart point ARGB color in C# | Excel pie chart color validation using Aspose.Cells | C# dictionary of expected colors for chart points

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

// The example loads an Excel workbook, accesses the first chart on the first worksheet, defines a dictionary of expected ARGB colors for each slice index, iterates through chart series and points, reads each slice's actual foreground color via ChartPoint.Area.ForegroundColor, compares it to the expected color, and outputs whether each slice matches or mismatches the predefined color.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the workbook containing the chart
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart on the worksheet (adjust index if needed)
            Chart chart = worksheet.Charts[0];

            // Define the expected RGB colors for each slice (index -> Color)
            Dictionary<int, Color> expectedColors = new Dictionary<int, Color>()
            {
                { 0, Color.FromArgb(255, 255, 0, 0) },   // Red
                { 1, Color.FromArgb(255, 0, 255, 0) },   // Green
                { 2, Color.FromArgb(255, 0, 0, 255) }    // Blue
                // Add more entries as required
            };

            // Iterate through each series in the chart
            foreach (Series series in chart.NSeries)
            {
                // Iterate through each point (slice) in the series
                for (int i = 0; i < series.Points.Count; i++)
                {
                    ChartPoint point = series.Points[i];

                    // Retrieve the actual foreground color of the slice via the Area object
                    Color actualColor = point.Area.ForegroundColor;

                    // Compare with the expected color if defined
                    if (expectedColors.TryGetValue(i, out Color expectedColor))
                    {
                        bool colorsMatch = actualColor.ToArgb() == expectedColor.ToArgb();
                        Console.WriteLine($"Slice {i}: {(colorsMatch ? "Match" : "Mismatch")} - Expected: {expectedColor}, Actual: {actualColor}");
                    }
                    else
                    {
                        Console.WriteLine($"Slice {i}: No expected color defined for this index.");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
