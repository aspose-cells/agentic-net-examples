// Title: Extract every chart from an Excel workbook and save them as SVG files with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export all worksheet charts to SVG files in a specific folder | How to batch convert Excel charts to SVG using Aspose.Cells .NET API | Save each chart from an Excel workbook as separate SVG images with Aspose.Cells | Aspose.Cells chart.ToImage SVG output for multiple charts in a workbook | Create subdirectory for chart SVG exports when processing an Excel file in C#
// Tags: Aspose.Cells chart export SVG | batch export Excel charts .NET | chart.ToImage SVG rendering | safe file names for chart images | output subfolder creation for chart exports

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace ExtractChartsToSvg
{
    // The C# program validates input, loads an Excel workbook with Aspose.Cells, creates a "ChartsSvg" subfolder (or a user‑specified folder), iterates through every worksheet and each chart, builds a filesystem‑safe filename using the sheet name and chart index, configures ImageOrPrintOptions for SVG output, and calls chart.ToImage to export each chart as an individual SVG file while logging successes and handling errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Validate arguments
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: ExtractChartsToSvg <workbookPath> [outputFolder]");
                return;
            }

            string workbookPath = args[0];
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook file not found at '{workbookPath}'.");
                return;
            }

            // Determine output folder (subfolder "ChartsSvg" under the workbook directory by default)
            string baseOutputFolder = args.Length > 1
                ? args[1]
                : Path.Combine(Path.GetDirectoryName(workbookPath) ?? string.Empty, "ChartsSvg");

            try
            {
                Directory.CreateDirectory(baseOutputFolder);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: Unable to create output folder '{baseOutputFolder}'. {ex.Message}");
                return;
            }

            try
            {
                // Load the workbook
                using (Workbook workbook = new Workbook(workbookPath))
                {
                    int totalChartsExported = 0;

                    // Iterate through each worksheet
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Iterate through each chart in the worksheet
                        for (int i = 0; i < sheet.Charts.Count; i++)
                        {
                            Chart chart = sheet.Charts[i];

                            // Prepare SVG export options
                            ImageOrPrintOptions options = new ImageOrPrintOptions
                            {
                                ImageType = ImageType.Svg
                                // SvgCssPrefix is obsolete; omitted
                            };

                            // Build a safe file name: SheetName_ChartIndex.svg
                            string safeSheetName = MakeFileSystemSafe(sheet.Name);
                            string svgFileName = $"{safeSheetName}_Chart{i + 1}.svg";
                            string svgFullPath = Path.Combine(baseOutputFolder, svgFileName);

                            try
                            {
                                // Export chart to SVG
                                chart.ToImage(svgFullPath, options);
                                Console.WriteLine($"Exported chart '{chart.Name}' from sheet '{sheet.Name}' to '{svgFullPath}'.");
                                totalChartsExported++;
                            }
                            catch (Exception exChart)
                            {
                                Console.WriteLine($"Failed to export chart '{chart.Name}' from sheet '{sheet.Name}': {exChart.Message}");
                            }
                        }
                    }

                    Console.WriteLine($"Finished. Total charts exported: {totalChartsExported}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
            }
        }

        // Replaces invalid filename characters with underscores
        private static string MakeFileSystemSafe(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }
            return name;
        }
    }
}
