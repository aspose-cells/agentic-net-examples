// Title: Convert all charts in multiple XLSX files to high‑resolution PNGs with Chinese localization using Aspose.Cells for .NET
// Common Searches: how to batch export Excel charts to PNG with Chinese labels using Aspose.Cells C# | Aspose.Cells render chart images from a folder of .xlsx files in .NET | apply custom GlobalizationSettings for Chinese error messages when converting charts with Aspose.Cells | C# code to save each worksheet chart as 300 DPI PNG using Aspose.Cells | sanitize worksheet names for file system when exporting chart images in C#
// Tags: Aspose.Cells batch chart to PNG conversion | Chinese GlobalizationSettings for Excel chart rendering | high‑resolution 300 DPI chart image export Aspose.Cells | safe file name generation for worksheet chart exports | iterate XLSX files in a directory C# Aspose.Cells | localize chart titles and axis labels Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;   // For ImageType enum

namespace ChartToLocalizedPng
{
    // Custom globalization settings that provide Chinese strings for boolean and error values.
    // The sample scans a given directory for .xlsx workbooks, applies a custom Chinese GlobalizationSettings to each workbook, optionally localizes chart titles and axis labels, and renders every chart on every worksheet to a 300 dpi PNG image using Aspose.Cells. Output files are named with the workbook and sanitized sheet names and saved to a target folder.
    public class ChineseGlobalizationSettings : GlobalizationSettings
    {
        public override string GetBooleanValueString(bool value)
        {
            // TRUE => "真", FALSE => "假"
            return value ? "真" : "假";
        }

        public override string GetErrorValueString(string err)
        {
            // Map common Excel error strings to Chinese equivalents.
            switch (err)
            {
                case "#NAME?": return "#名称?";
                case "#DIV/0!": return "#除以0!";
                case "#REF!": return "#引用!";
                case "#VALUE!": return "#值!";
                case "#N/A": return "#无效!";
                case "#NUM!": return "#数字!";
                case "#NULL!": return "#空!";
                default: return base.GetErrorValueString(err);
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Resolve input and output folders.
            string inputFolder;
            string outputFolder;

            if (args.Length >= 2)
            {
                inputFolder = args[0];
                outputFolder = args[1];
            }
            else
            {
                Console.WriteLine("Usage: ChartToLocalizedPng <inputFolder> <outputFolder>");
                Console.WriteLine("No arguments supplied – using current directory for both input and output.");
                inputFolder = Directory.GetCurrentDirectory();
                outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "ChartImages");
            }

            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            Directory.CreateDirectory(outputFolder);

            string[] xlsxFiles = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);
            if (xlsxFiles.Length == 0)
            {
                Console.WriteLine($"No .xlsx files found in {inputFolder}");
                return;
            }

            foreach (string filePath in xlsxFiles)
            {
                // Verify the file actually exists before attempting to load.
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found (skipped): {filePath}");
                    continue;
                }

                try
                {
                    Console.WriteLine($"Processing workbook: {Path.GetFileName(filePath)}");
                    Workbook workbook = new Workbook(filePath);

                    // Apply Chinese globalization settings to the whole workbook.
                    workbook.Settings.GlobalizationSettings = new ChineseGlobalizationSettings();

                    // Iterate through worksheets.
                    for (int wsIndex = 0; wsIndex < workbook.Worksheets.Count; wsIndex++)
                    {
                        Worksheet sheet = workbook.Worksheets[wsIndex];
                        ChartCollection charts = sheet.Charts;

                        if (charts == null || charts.Count == 0)
                            continue; // No charts on this sheet.

                        for (int chartIndex = 0; chartIndex < charts.Count; chartIndex++)
                        {
                            try
                            {
                                Chart chart = charts[chartIndex];

                                // Localize chart title and axis titles (optional).
                                if (chart.Title != null)
                                    chart.Title.Text = "图表标题";

                                if (chart.ValueAxis?.Title != null)
                                    chart.ValueAxis.Title.Text = "数值轴";

                                if (chart.CategoryAxis?.Title != null)
                                    chart.CategoryAxis.Title.Text = "类别轴";

                                // Prepare image options – PNG with high resolution.
                                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                                {
                                    ImageType = ImageType.Png,
                                    HorizontalResolution = 300,
                                    VerticalResolution = 300,
                                    Transparent = false,
                                    OnePagePerSheet = true
                                };

                                // Build a safe output file name.
                                string safeSheetName = MakeFileSystemSafe(sheet.Name);
                                string outputFileName = $"{Path.GetFileNameWithoutExtension(filePath)}_{safeSheetName}_Chart{chartIndex + 1}.png";
                                string outputPath = Path.Combine(outputFolder, outputFileName);

                                // Render chart to PNG file.
                                chart.ToImage(outputPath, imgOptions);
                                Console.WriteLine($"  Saved chart {chartIndex + 1} of sheet \"{sheet.Name}\" to {outputFileName}");
                            }
                            catch (Exception chartEx)
                            {
                                Console.WriteLine($"  Error rendering chart {chartIndex + 1} on sheet \"{sheet.Name}\": {chartEx.Message}");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file {Path.GetFileName(filePath)}: {ex.Message}");
                }
            }

            Console.WriteLine("Processing completed.");
        }

        // Helper method to replace characters that are invalid in file names.
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
