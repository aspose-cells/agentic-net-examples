// Title: Apply a custom color palette to every chart series in an XLS workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an existing XLS file with Aspose.Cells, creates an array of System.Drawing.Color values, and assigns each color to the fill, border, and marker of every chart series across all worksheets before saving the workbook. | Demonstrate how to loop through all worksheets and charts in Aspose.Cells, using the NSeries collection to set series.Area.ForegroundColor, series.Border.Color, and series.Marker.ForegroundColor from a repeating custom palette.
// Common Searches: Aspose.Cells C# change colors of all chart series in an existing XLS workbook | how to set a custom palette for chart series programmatically in .NET using Aspose.Cells | apply the same color scheme to multiple charts in an XLS file with Aspose.Cells | C# iterate through worksheets and charts to modify series colors in Aspose.Cells | save modified XLS workbook after updating chart series colors Aspose.Cells
// Tags: apply custom chart series palette Aspose.Cells | iterate worksheets charts Aspose.Cells C# | set series fill border marker color Aspose.Cells | custom RGB colors for chart series Aspose.Cells | save XLS workbook after chart modification Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing XLS workbook (or creates a new one), defines a custom System.Drawing.Color array, iterates through every worksheet and each chart, and applies the colors to each series' area fill, border, and marker. Finally, it saves the updated workbook to a new XLS file.
class Program
{
    static void Main()
    {
        try
        {
            // Input workbook path (replace with actual file path if needed)
            string inputPath = "input.xls";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                if (workbook.Worksheets.Count > 0)
                {
                    workbook.Worksheets[0].Name = "Sheet1";
                }
            }

            // Define a custom color palette (example colors)
            Color[] customPalette = new Color[]
            {
                Color.FromArgb(0xFF, 0x4E, 0x79, 0xA7), // Blue
                Color.FromArgb(0xFF, 0xC0, 0x50, 0x4D), // Red
                Color.FromArgb(0xFF, 0x9B, 0xC9, 0x5A), // Green
                Color.FromArgb(0xFF, 0xFF, 0xC0, 0x00), // Yellow
                Color.FromArgb(0xFF, 0x8E, 0x44, 0xAD)  // Purple
            };

            // Apply custom colors to chart series in all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    try
                    {
                        var series = chart.NSeries;
                        for (int i = 0; i < series.Count; i++)
                        {
                            Color paletteColor = customPalette[i % customPalette.Length];

                            // Set fill color for the series area (works for column/bar charts)
                            series[i].Area.ForegroundColor = paletteColor;

                            // Set line color (works for line/area charts) via Border
                            series[i].Border.Color = paletteColor;

                            // Set marker color (works for scatter/line charts)
                            series[i].Marker.ForegroundColor = paletteColor;
                        }
                    }
                    catch (Exception exChart)
                    {
                        Console.WriteLine($"Warning: Failed to apply colors to chart '{chart.Name}'. {exChart.Message}");
                    }
                }
            }

            // Output workbook path (replace with desired output path)
            string outputPath = "output.xls";

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
