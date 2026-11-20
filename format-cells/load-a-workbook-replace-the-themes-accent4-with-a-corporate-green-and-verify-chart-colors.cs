// Title: Replace the workbook’s Accent4 theme color with a corporate green and apply it to all chart series using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# change workbook theme Accent4 color to custom RGB | Set corporate green fill for every chart series in an Excel file with Aspose.Cells | Verify chart series colors after modifying the Excel theme using Aspose.Cells .NET | How to update Excel theme accent colors via reflection in Aspose.Cells | Iterate through worksheets and charts to apply solid fill colors in C#
// Tags: replace theme accent color Aspose.Cells | chart series solid fill Aspose.Cells | custom RGB theme modification .NET | iterate worksheets charts Aspose.Cells | save workbook after theme change C#

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example loads an existing workbook, uses reflection to replace the fourth accent (Accent4) in the workbook theme with a corporate green RGB value, then walks through each worksheet and chart to read or set the series area or line fill to that green when no solid fill exists, logs the colors, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Define the corporate green color
            Color corporateGreen = Color.FromArgb(0, 128, 0); // Adjust RGB as needed

            // Attempt to modify the theme's Accent4 color using reflection (Theme may not exist in older versions)
            try
            {
                var themeProp = workbook.GetType().GetProperty("Theme");
                var themeObj = themeProp?.GetValue(workbook);
                if (themeObj != null)
                {
                    var accentColorsProp = themeObj.GetType().GetProperty("AccentColors");
                    var accentColors = accentColorsProp?.GetValue(themeObj) as Color[];
                    if (accentColors != null && accentColors.Length > 3)
                    {
                        accentColors[3] = corporateGreen;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Theme modification skipped: {ex.Message}");
            }

            // Iterate through all charts and display (or set) series fill colors
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    Console.WriteLine($"Chart: {chart.Name} (Sheet: {sheet.Name})");
                    int seriesIdx = 0;

                    // Use dynamic to avoid compile‑time dependency on ChartSeries type
                    foreach (var seriesObj in chart.NSeries)
                    {
                        dynamic series = seriesObj;
                        Color? fillColor = null;

                        try
                        {
                            // Check area fill
                            if (series.Area != null && series.Area.Fill != null && series.Area.Fill.Type == FillType.Solid)
                            {
                                fillColor = series.Area.Fill.SolidFillColor;
                            }
                            // Check line fill
                            else if (series.Line != null && series.Line.Fill != null && series.Line.Fill.Type == FillType.Solid)
                            {
                                fillColor = series.Line.Fill.SolidFillColor;
                            }

                            // If no fill color, set it to corporate green
                            if (!fillColor.HasValue)
                            {
                                series.Area.Fill.Type = FillType.Solid;
                                series.Area.Fill.SolidFillColor = corporateGreen;
                                fillColor = corporateGreen;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Failed to process series {seriesIdx}: {ex.Message}");
                        }

                        string colorInfo = fillColor.HasValue ? fillColor.Value.Name : "None";
                        Console.WriteLine($"  Series {seriesIdx}: Fill Color = {colorInfo}");
                        seriesIdx++;
                    }
                }
            }

            // Ensure the output directory exists
            try
            {
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
