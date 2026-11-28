// Title: Export each worksheet's charts to high‑resolution PNG files named after the sheet using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export each worksheet chart to separate PNG file with sheet name | How to save Excel chart images as 300 dpi PNG using Aspose.Cells .NET | Generate PNG images for all charts in a workbook and include worksheet name in filename | Aspose.Cells chart.ToImage high resolution PNG per worksheet example | Command‑line tool to extract charts from Excel to PNG with Aspose.Cells
// Tags: chart export to PNG Aspose.Cells | worksheet iteration for chart extraction .NET | high‑resolution chart image options C# | safe filename generation for Excel sheet names | batch chart image generation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing; // For ImageType enum

namespace AsposeCellsChartExport
{
    // The console application validates command‑line arguments, loads an Excel workbook with Aspose.Cells, creates the output folder, and iterates through every worksheet. For each chart it builds a safe filename that combines the worksheet name and chart index, configures ImageOrPrintOptions for 300 dpi PNG output, and calls chart.ToImage to save the chart as a high‑resolution PNG file. Progress and errors are logged, and the workbook is disposed after processing.
    class Program
    {
        static void Main(string[] args)
        {
            // Validate input arguments
            if (args.Length < 1)
            {
                Console.WriteLine("Usage: AsposeCellsChartExport <input_excel_file_path> [output_directory]");
                return;
            }

            string inputPath = args[0];
            string outputDir = args.Length >= 2 ? args[1] : Path.GetDirectoryName(inputPath) ?? Directory.GetCurrentDirectory();

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" does not exist.");
                return;
            }

            Workbook workbook = null;
            try
            {
                // Load the workbook
                workbook = new Workbook(inputPath);
                Console.WriteLine($"Workbook \"{Path.GetFileName(inputPath)}\" loaded successfully.");

                // Ensure output directory exists
                Directory.CreateDirectory(outputDir);

                // Iterate through each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    string sheetName = sheet.Name;
                    int chartCount = sheet.Charts.Count;

                    if (chartCount == 0)
                    {
                        Console.WriteLine($"Worksheet \"{sheetName}\" contains no charts. Skipping.");
                        continue;
                    }

                    Console.WriteLine($"Worksheet \"{sheetName}\" contains {chartCount} chart(s).");

                    // Iterate through each chart in the worksheet
                    for (int i = 0; i < chartCount; i++)
                    {
                        try
                        {
                            Chart chart = sheet.Charts[i];

                            // Build a safe file name that includes the worksheet name and chart index
                            string safeSheetName = string.Concat(sheetName.Split(Path.GetInvalidFileNameChars()));
                            string outputFileName = $"{safeSheetName}_Chart{i + 1}.png";
                            string outputPath = Path.Combine(outputDir, outputFileName);

                            // Set image options (high resolution PNG)
                            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                            {
                                ImageType = ImageType.Png,
                                HorizontalResolution = 300,
                                VerticalResolution = 300
                            };

                            // Export the chart to PNG
                            chart.ToImage(outputPath, imgOptions);
                            Console.WriteLine($"Saved chart {i + 1} of worksheet \"{sheetName}\" to \"{outputPath}\".");
                        }
                        catch (Exception chartEx)
                        {
                            Console.WriteLine($"Failed to export chart {i + 1} on worksheet \"{sheetName}\": {chartEx.Message}");
                        }
                    }
                }

                Console.WriteLine("Processing completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
            finally
            {
                // Clean up resources
                workbook?.Dispose();
            }
        }
    }
}
