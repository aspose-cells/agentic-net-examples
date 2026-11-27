// Title: Export Excel charts to PNG for multiple locales in parallel while keeping memory usage low with Aspose.Cells for .NET
// Common Searches: how to export Excel chart images for each locale using Aspose.Cells in C# parallel processing | C# Aspose.Cells parallel chart export memory leak prevention | stable memory usage when generating localized chart PNGs with Aspose.Cells | multi‑threaded workbook chart extraction per locale Aspose.Cells .NET | release native resources after chart.ToImage in Aspose.Cells C#
// Tags: parallel chart image export Aspose.Cells | localized chart PNG generation C# | memory management Aspose.Cells chart processing | chart.ToImage native resource release | region-specific workbook calculation Aspose.Cells

using System;
using System.IO;
using System.Collections.Generic;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace LocalizedChartExport
{
    // The sample loads .xlsx workbooks from a folder, parses a comma‑separated list of locale identifiers, and uses Parallel.ForEach to process each workbook concurrently. For every locale it creates a dedicated output subfolder, sets Workbook.Settings.Region, recalculates formulas, iterates through worksheets and their charts, forces full chart calculation, and saves each chart as a PNG via Chart.ToImage. After each locale the code forces garbage collection to release native resources, ensuring stable memory consumption during multi‑threaded localized chart exports.
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 3)
            {
                Console.WriteLine("Usage: LocalizedChartExport <inputFolder> <outputFolder> <locales>");
                return;
            }

            string inputFolder = args[0];
            string outputFolder = args[1];
            string[] localeTokens = args[2].Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            Directory.CreateDirectory(outputFolder);

            // Parse locale strings into CountryCode enum values
            List<CountryCode> locales = new List<CountryCode>();
            foreach (var token in localeTokens)
            {
                if (Enum.TryParse(token.Trim(), true, out CountryCode code))
                {
                    locales.Add(code);
                }
                else
                {
                    Console.WriteLine($"Warning: Unable to parse locale '{token}'. Skipping.");
                }
            }

            if (locales.Count == 0)
            {
                Console.WriteLine("No valid locales supplied. Exiting.");
                return;
            }

            // Gather all .xlsx files
            string[] files = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);
            if (files.Length == 0)
            {
                Console.WriteLine("No .xlsx files found in the input folder.");
                return;
            }

            Console.WriteLine($"Found {files.Length} workbook(s). Starting export with {locales.Count} locale(s).");

            var parallelOptions = new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount };

            Parallel.ForEach(files, parallelOptions, filePath =>
            {
                try
                {
                    ExportWorkbookCharts(filePath, outputFolder, locales);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{Path.GetFileName(filePath)}': {ex.Message}");
                }
            });

            Console.WriteLine("Export completed.");
        }

        private static void ExportWorkbookCharts(string workbookPath, string outputRoot, List<CountryCode> locales)
        {
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook file not found: {workbookPath}");
                return;
            }

            string workbookName = Path.GetFileNameWithoutExtension(workbookPath);

            foreach (CountryCode locale in locales)
            {
                // Each locale gets its own folder
                string localeFolder = Path.Combine(outputRoot, $"{workbookName}_{locale}");
                Directory.CreateDirectory(localeFolder);

                try
                {
                    using (Workbook workbook = new Workbook(workbookPath))
                    {
                        // Apply locale for number/date formatting
                        workbook.Settings.Region = locale;

                        // Recalculate all formulas under the current locale
                        workbook.CalculateFormula();

                        // Process each worksheet
                        foreach (Worksheet sheet in workbook.Worksheets)
                        {
                            // Process each chart in the worksheet
                            foreach (Chart chart in sheet.Charts)
                            {
                                try
                                {
                                    // Ensure chart data points are fully calculated (important for localized values)
                                    ChartCalculateOptions calcOptions = new ChartCalculateOptions
                                    {
                                        UpdateAllPoints = true
                                    };
                                    chart.Calculate(calcOptions);

                                    // Build a deterministic file name
                                    int chartIdx = sheet.Charts.IndexOf(chart);
                                    string chartName = !string.IsNullOrEmpty(chart.Name) ? chart.Name : $"Chart{chartIdx}";
                                    string chartFileName = $"{sheet.Name}_{chartName}_{locale}.png";
                                    string chartPath = Path.Combine(localeFolder, chartFileName);

                                    // Directly save the chart image to file (avoids System.Drawing dependency)
                                    chart.ToImage(chartPath);
                                }
                                catch (Exception exChart)
                                {
                                    Console.WriteLine($"Error processing chart in sheet '{sheet.Name}': {exChart.Message}");
                                }
                            }
                        }
                    }
                }
                catch (Exception exWorkbook)
                {
                    Console.WriteLine($"Error loading workbook '{Path.GetFileName(workbookPath)}': {exWorkbook.Message}");
                }

                // Force garbage collection to release native resources after each locale
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}
