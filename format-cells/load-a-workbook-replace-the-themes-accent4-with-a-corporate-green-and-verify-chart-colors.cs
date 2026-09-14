// Title: How to replace the Accent4 theme color with a corporate green in an Excel workbook and verify chart series colors using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an existing .xlsx file, changes the workbook's Accent4 theme color to RGB #00A651 with Aspose.Cells, and saves the modified file. | Write a C# loop that enumerates every chart in a workbook and prints each series' fill or border color after the theme update. | Create a C# snippet that checks for the input file, creates the output directory if it does not exist, and handles exceptions while applying the theme change.
// Common Searches: Aspose.Cells change workbook theme Accent4 to custom RGB color in C# | C# iterate over Excel charts and read series fill color using Aspose.Cells | verify chart series colors after modifying Excel theme with Aspose.Cells .NET | ensure output folder exists when saving a modified workbook with Aspose.Cells
// Tags: set theme accent4 color Aspose.Cells | custom corporate green Excel theme .NET | read chart series fill color Aspose.Cells | validate workbook file existence C# | save workbook with updated theme Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.Drawing;
using System.IO;

// The example loads an existing Excel workbook, updates the Accent4 theme color to a corporate green (RGB #00A651), iterates through all worksheets and charts to output each series' fill or border color, ensures the output directory exists, and saves the workbook with the new theme.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Define the corporate green color (hex #00A651)
            Color corporateGreen = Color.FromArgb(0, 166, 81);

            // Change the theme's Accent4 color if the workbook supports themes
            try
            {
                workbook.SetThemeColor(ThemeColorType.Accent4, corporateGreen);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unable to modify theme color: {ex.Message}");
            }

            // Verify chart colors after the theme change
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    Console.WriteLine($"Chart \"{chart.Name}\" in worksheet \"{sheet.Name}\":");

                    int seriesIndex = 0;
                    foreach (Series series in chart.NSeries)
                    {
                        // Attempt to retrieve the series fill color
                        Color? seriesColor = null;

                        // Most chart types store fill color in the Area's ForegroundColor
                        if (series.Area != null && series.Area.ForegroundColor != Color.Empty)
                        {
                            seriesColor = series.Area.ForegroundColor;
                        }
                        // Fallback to border color if area color is not set
                        else if (series.Border != null && series.Border.Color != Color.Empty)
                        {
                            seriesColor = series.Border.Color;
                        }

                        string colorInfo = seriesColor.HasValue ? seriesColor.Value.Name : "Not set";
                        Console.WriteLine($"  Series {seriesIndex} color: {colorInfo}");
                        seriesIndex++;
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
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
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
