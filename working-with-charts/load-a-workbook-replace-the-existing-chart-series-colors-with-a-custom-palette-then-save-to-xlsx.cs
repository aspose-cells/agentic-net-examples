// Title: Replace chart series colors with a custom palette in an existing XLSX workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# change colors of all chart series in an existing workbook | apply custom color palette to Excel chart series with Aspose.Cells .NET | iterate worksheets and charts to set series fill and border colors in C# | save modified workbook after updating chart series colors using Aspose.Cells
// Tags: chart series color customization Aspose.Cells | custom palette for Excel chart series .NET | iterate workbook charts Aspose.Cells API | set series area foreground color Aspose.Cells | save workbook as XLSX after chart styling Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing XLSX file, defines a custom color array, loops through every worksheet and chart, assigns each series a palette color for both fill and border, and saves the updated workbook as a new XLSX file.
class ReplaceChartSeriesColors
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Define a custom color palette
            Color[] customPalette = new Color[]
            {
                Color.FromArgb(0xFF, 0x4E, 0x79, 0xA7), // Blue
                Color.FromArgb(0xFF, 0xA5, 0xA5, 0xA5), // Gray
                Color.FromArgb(0xFF, 0xFF, 0xC0, 0x00), // Gold
                Color.FromArgb(0xFF, 0x9F, 0x00, 0xA7), // Purple
                Color.FromArgb(0xFF, 0x00, 0xA1, 0xE0)  // Light Blue
            };

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all charts in the worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    try
                    {
                        // Access the series collection
                        var seriesCollection = chart.NSeries;

                        for (int i = 0; i < seriesCollection.Count; i++)
                        {
                            Series series = seriesCollection[i];

                            // Choose a color from the custom palette (cycle if more series than colors)
                            Color paletteColor = customPalette[i % customPalette.Length];

                            // Apply the color to the series area (fill) and border (line)
                            series.Area.ForegroundColor = paletteColor;
                            series.Border.Color = paletteColor;
                        }
                    }
                    catch (Exception exChart)
                    {
                        Console.WriteLine($"Error processing chart on sheet '{sheet.Name}': {exChart.Message}");
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook to a new file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
