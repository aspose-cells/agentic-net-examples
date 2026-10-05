// Title: Replace chart series colors with a custom rotating palette in an XLSX workbook using Aspose.Cells for .NET
// AI Prompts: Load an existing XLSX workbook, loop through each worksheet and chart, and assign colors from a predefined Color[] array to the series area, line, and marker before saving the file. | Use Aspose.Cells C# API to apply a cyclic custom palette to every chart series in a workbook, updating fill, border, and marker colors, then export the workbook as XLSX.
// Common Searches: asp.net change all chart series colors programmatically with Aspose.Cells | c# apply custom palette to Excel chart series using Aspose.Cells | replace chart series fill and line colors in every worksheet Aspose.Cells example | how to set marker color for chart series in Aspose.Cells .NET
// Tags: chart series color assignment Aspose.Cells | iterate workbook charts C# | set series fill and line colors .NET | export workbook to XLSX after chart styling | cyclic palette implementation for Excel charts

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program loads input.xlsx, defines a five‑color palette, iterates through every worksheet and chart, assigns each series a palette color (cycling if needed) for its area fill, line, and marker, and then saves the updated workbook to output.xlsx.
class ReplaceChartSeriesColors
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Define a custom palette of colors
            Color[] customPalette = new Color[]
            {
                Color.FromArgb(0x4E, 0x79, 0xA7), // Blue
                Color.FromArgb(0xF2, 0xC9, 0x4D), // Yellow
                Color.FromArgb(0x9A, 0xC9, 0x5B), // Green
                Color.FromArgb(0xE5, 0x7A, 0x5A), // Red
                Color.FromArgb(0xB5, 0x7F, 0xA5)  // Purple
            };

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all charts in the worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    // Iterate through all series in the chart
                    for (int i = 0; i < chart.NSeries.Count; i++)
                    {
                        // Use dynamic to avoid compile‑time dependency on ChartSeries type
                        dynamic series = chart.NSeries[i];

                        // Choose a color from the custom palette (cycle if more series than colors)
                        Color paletteColor = customPalette[i % customPalette.Length];

                        // Apply the color to the series fill (area) and line
                        series.Area.ForegroundColor = paletteColor;   // Fill color
                        series.Line.Color = paletteColor;             // Border/line color

                        // Optionally set marker color if markers are used
                        series.Marker.ForegroundColor = paletteColor;
                    }
                }
            }

            // Save the modified workbook to a new XLSX file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
