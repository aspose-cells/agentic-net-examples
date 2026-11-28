// Title: Export every chart in an Excel workbook or a folder of workbooks to separate PNG files using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export all worksheet charts to PNG files in a batch | How to save each chart from an .xlsx workbook as an individual PNG image with Aspose.Cells | Process a directory of Excel files and extract charts as PNG images using Aspose.Cells .NET | Aspose.Cells chart.ToImage PNG export example for multiple workbooks | Create separate output folders for chart PNGs when converting Excel charts in C#
// Tags: export charts to PNG Aspose.Cells | batch chart extraction from Excel workbooks | chart.ToImage ImageOrPrintOptions usage | per‑workbook chart export directories | process multiple .xlsx files Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;   // Required for ImageType enum

namespace ExportChartsToPng
{
    // A C# console utility that accepts a single workbook or a folder of .xlsx files, creates a dedicated output directory for each workbook, loads each file with Aspose.Cells, and iterates through all worksheets. For every chart it generates a safe filename based on the sheet name and chart index, then exports the chart to a PNG file using ImageOrPrintOptions. The program logs each successful export, handles loading and export errors, and reports the total number of charts exported per workbook.
    class Program
    {
        /// <param name="args">
        /// args[0] – Path to a workbook file (e.g., "sample.xlsx") or a folder containing workbooks.
        /// Optional args[1] – Output directory. If omitted, a sub‑folder named "ChartExports" is created next to each source file.
        /// </param>
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: ExportChartsToPng <workbookPath|folderPath> [outputFolder]");
                return;
            }

            string inputPath = args[0];
            string? outputRoot = args.Length > 1 ? args[1] : null;

            try
            {
                if (File.Exists(inputPath))
                {
                    ProcessWorkbook(inputPath, outputRoot);
                }
                else if (Directory.Exists(inputPath))
                {
                    string[] files = Directory.GetFiles(inputPath, "*.xlsx", SearchOption.TopDirectoryOnly);
                    if (files.Length == 0)
                    {
                        Console.WriteLine($"No .xlsx files found in folder: {inputPath}");
                        return;
                    }

                    foreach (string file in files)
                    {
                        ProcessWorkbook(file, outputRoot);
                    }
                }
                else
                {
                    Console.WriteLine($"The specified path does not exist: {inputPath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }

        /// <param name="workbookPath">Full path to the source workbook.</param>
        /// <param name="outputRoot">
        /// Optional root folder for exported images. If null, a folder named "ChartExports"
        /// will be created beside the source workbook.
        /// </param>
        private static void ProcessWorkbook(string workbookPath, string? outputRoot)
        {
            Console.WriteLine($"Processing workbook: {workbookPath}");

            // Determine output directory
            string baseOutputDir = outputRoot ??
                                   Path.Combine(Path.GetDirectoryName(workbookPath) ?? string.Empty, "ChartExports");

            // Create a sub‑folder for this workbook to avoid name clashes
            string workbookNameWithoutExt = Path.GetFileNameWithoutExtension(workbookPath);
            string workbookOutputDir = Path.Combine(baseOutputDir, workbookNameWithoutExt);
            Directory.CreateDirectory(workbookOutputDir);

            // Load the workbook
            Workbook wb;
            try
            {
                wb = new Workbook(workbookPath);
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load workbook '{workbookPath}': {loadEx.Message}");
                return;
            }

            // Prepare image options for PNG output
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png   // Export charts as PNG
            };

            int totalChartsExported = 0;

            // Iterate through all worksheets
            foreach (Worksheet sheet in wb.Worksheets)
            {
                int chartCount = sheet.Charts.Count;
                if (chartCount == 0)
                {
                    continue; // No charts on this sheet
                }

                for (int i = 0; i < chartCount; i++)
                {
                    Chart chart = sheet.Charts[i];

                    // Build a unique file name: SheetName_ChartIndex.png
                    string safeSheetName = MakeFileSystemSafe(sheet.Name);
                    string fileName = $"{safeSheetName}_Chart{i + 1}.png";
                    string fullPath = Path.Combine(workbookOutputDir, fileName);

                    try
                    {
                        // Export the chart to PNG
                        chart.ToImage(fullPath, imgOptions);
                        Console.WriteLine($"Exported chart to: {fullPath}");
                        totalChartsExported++;
                    }
                    catch (Exception chartEx)
                    {
                        Console.WriteLine($"Failed to export chart (Sheet: {sheet.Name}, Index: {i}): {chartEx.Message}");
                    }
                }
            }

            Console.WriteLine($"Finished processing '{workbookPath}'. Total charts exported: {totalChartsExported}");
        }

        // Replace invalid filename characters with underscores
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
