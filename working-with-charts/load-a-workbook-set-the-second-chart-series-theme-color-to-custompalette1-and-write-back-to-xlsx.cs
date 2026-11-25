// Title: Change the second chart series to the workbook’s Accent1 theme color and save the workbook as XLSX using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells set second chart series fill to workbook Accent1 theme color | How to apply a custom palette color to a specific series in an Excel chart with Aspose.Cells | Use GetThemeColor to change chart series border color in Aspose.Cells .NET | Save an updated Excel workbook after modifying chart series colors using Aspose.Cells | Aspose.Cells example for changing chart series theme color and exporting to XLSX | Modify second series area fill format in Excel chart programmatically with Aspose.Cells
// Tags: Aspose.Cells set chart series theme color | apply workbook Accent1 color to chart series C# | modify second series fill format Aspose.Cells | chart series border color using GetThemeColor | save workbook as XLSX after chart update Aspose.Cells | Excel chart series area fill pattern Aspose.Cells | use GetThemeColor Accent1 Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartSeriesTheme
{
    // The example loads an existing XLSX workbook, accesses the first worksheet's first chart, changes the fill and border colors of the second series to the workbook's Accent1 theme color (used as a custom palette entry), and then saves the modified workbook as a new XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to the input and output files.
            // Adjust these paths as needed for your environment.
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                return;
            }

            try
            {
                // Load the workbook from the specified file.
                Workbook workbook = new Workbook(inputPath);

                // Assume we are working with the first worksheet.
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure that the worksheet contains at least one chart.
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("Error: No charts found in the first worksheet.");
                    return;
                }

                // Get the first chart in the worksheet.
                Chart chart = worksheet.Charts[0];

                // Ensure the chart has at least two series.
                SeriesCollection seriesCollection = chart.NSeries;
                if (seriesCollection.Count < 2)
                {
                    Console.WriteLine("Error: The chart does not contain a second series.");
                    return;
                }

                // Access the second series (index 1).
                Series secondSeries = seriesCollection[1];

                // ------------------------------------------------------------
                // Set the theme color for the second series.
                // Aspose.Cells does not expose a direct "CustomPalette1" enum,
                // but the typical way to apply a theme‑based color is to use
                // the workbook's GetThemeColor method with a ThemeColorType.
                // Here we map "CustomPalette1" to the Accent1 theme color,
                // which is commonly the first custom palette entry.
                // ------------------------------------------------------------
                Color themeColor = workbook.GetThemeColor(ThemeColorType.Accent1);

                // Apply the theme color to the series area.
                secondSeries.Area.FillFormat.Pattern = FillPattern.Solid;
                secondSeries.Area.ForegroundColor = themeColor;

                // Optionally, set the border color to match the theme.
                secondSeries.Border.Color = themeColor;

                // Ensure the output directory exists.
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook back to XLSX format.
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An exception occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
