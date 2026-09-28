// Title: How to assign distinct theme colors to each series in an Excel chart using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that iterates over a chart's SeriesCollection and sets a unique theme color for each series based on its index. | Update an existing Aspose.Cells workbook to apply both fill (foreground) and border colors to every chart series using a predefined Color[] palette.
// Common Searches: C# Aspose.Cells set different colors for each series in an Excel chart | How to loop through chart series and apply theme colors with Aspose.Cells .NET | Assign foreground and border colors to Excel chart series programmatically using Aspose.Cells | Cycle through a color palette when styling chart series in Aspose.Cells C# | Update first chart in a workbook to use custom series colors with Aspose.Cells
// Tags: Aspose.Cells chart series color styling | C# set series foreground color Aspose.Cells | apply border color to Excel chart series .NET | iterate chart NSeries collection Aspose.Cells | custom theme palette for Excel chart series C#

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads a workbook, accesses the first worksheet's first chart, defines a Color[] of theme colors, iterates through each series in the chart, and assigns each series a foreground and border color from the palette (cycling when necessary), then saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
                throw new InvalidOperationException("No charts found in the first worksheet.");

            // Get the first chart on the worksheet
            Chart chart = sheet.Charts[0];

            // Ensure the chart has at least one series
            if (chart.NSeries.Count == 0)
                throw new InvalidOperationException("The chart does not contain any series.");

            // Define a set of theme colors to apply to series
            Color[] themeColors = new Color[]
            {
                Color.FromArgb(0xFF, 0x5B, 0x9B, 0xD5), // Theme color 1
                Color.FromArgb(0xFF, 0xED, 0x7D, 0x31), // Theme color 2
                Color.FromArgb(0xFF, 0xA5, 0xA5, 0xA5), // Theme color 3
                Color.FromArgb(0xFF, 0x70, 0xAD, 0x47)  // Theme color 4
            };

            // Iterate through each series in the chart and assign a theme color based on its index
            for (int i = 0; i < chart.NSeries.Count; i++)
            {
                Series series = chart.NSeries[i];
                Color color = themeColors[i % themeColors.Length];

                // Apply the color to the series (fill/line color)
                series.Area.ForegroundColor = color;
                series.Border.Color = color; // optional: set border color as well
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
