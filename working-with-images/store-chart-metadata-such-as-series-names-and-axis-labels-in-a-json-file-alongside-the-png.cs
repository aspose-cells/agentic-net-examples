// Title: Export Excel charts to PNG and save chart series names and axis titles as JSON using Aspose.Cells for .NET
// Common Searches: how to render an Excel chart as PNG and export its series names to JSON with Aspose.Cells C# | Aspose.Cells extract chart series and axis titles and write them to a JSON file | C# generate PNG images of worksheet pages and store chart metadata alongside them | save chart type, series list and axis labels from a workbook to JSON while exporting images | iterate worksheets with Aspose.Cells, render charts to PNG and capture chart details in JSON
// Tags: export chart to PNG Aspose.Cells .NET | serialize chart metadata to JSON C# | extract chart series names Aspose.Cells | retrieve chart axis titles Aspose.Cells | render worksheet pages as images Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Charts;

namespace ChartExportWithMetadata
{
    // The sample loads an Excel workbook with Aspose.Cells, iterates each worksheet, renders every page to a PNG file, and for each chart extracts its type, series names, and axis titles. It then builds a JSON object containing workbook, worksheet, chart index, chart type, series names, and axis titles, and writes this metadata to a JSON file placed alongside the corresponding PNG image.
    class Program
    {
        static void Main(string[] args)
        {
            // Validate arguments
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: ChartExportWithMetadata <input_excel_file> [output_folder]");
                return;
            }

            string inputPath = args[0];
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                return;
            }

            string outputFolder = args.Length > 1 ? args[1] : Path.GetDirectoryName(Path.GetFullPath(inputPath));
            if (string.IsNullOrEmpty(outputFolder))
            {
                Console.WriteLine("Error: Unable to determine output folder.");
                return;
            }

            if (!Directory.Exists(outputFolder))
            {
                try
                {
                    Directory.CreateDirectory(outputFolder);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error: Unable to create output folder \"{outputFolder}\". {ex.Message}");
                    return;
                }
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Console.WriteLine($"Workbook \"{Path.GetFileName(inputPath)}\" loaded successfully.");

                // Iterate through all worksheets
                for (int wsIndex = 0; wsIndex < workbook.Worksheets.Count; wsIndex++)
                {
                    Worksheet sheet = workbook.Worksheets[wsIndex];
                    if (sheet.Charts.Count == 0)
                        continue; // No charts on this sheet

                    Console.WriteLine($"Processing worksheet \"{sheet.Name}\" with {sheet.Charts.Count} chart(s).");

                    // Prepare rendering options (default PNG)
                    ImageOrPrintOptions renderOptions = new ImageOrPrintOptions();

                    // Render each page of the worksheet to PNG
                    SheetRender sheetRender = new SheetRender(sheet, renderOptions);
                    int pageCount = sheetRender.PageCount;
                    for (int page = 0; page < pageCount; page++)
                    {
                        string pngFileName = $"{Path.GetFileNameWithoutExtension(inputPath)}_Sheet{wsIndex + 1}_Page{page + 1}.png";
                        string pngPath = Path.Combine(outputFolder, pngFileName);
                        try
                        {
                            sheetRender.ToImage(page, pngPath);
                            Console.WriteLine($"Rendered page {page + 1} of worksheet \"{sheet.Name}\" to \"{pngFileName}\".");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error rendering page {page + 1} of sheet \"{sheet.Name}\": {ex.Message}");
                        }
                    }

                    // Export metadata for each chart on the worksheet
                    for (int chartIdx = 0; chartIdx < sheet.Charts.Count; chartIdx++)
                    {
                        try
                        {
                            Chart chart = sheet.Charts[chartIdx];

                            // Gather series names
                            List<string> seriesNames = new List<string>();
                            for (int s = 0; s < chart.NSeries.Count; s++)
                            {
                                string name = chart.NSeries[s].Name;
                                // If the name is empty, keep it as empty string (no alternative property available)
                                seriesNames.Add(name ?? string.Empty);
                            }

                            // Gather axis titles (if any)
                            string valueAxisTitle = chart.ValueAxis?.Title?.Text ?? string.Empty;
                            string categoryAxisTitle = chart.CategoryAxis?.Title?.Text ?? string.Empty;

                            // Build metadata object
                            var metadata = new
                            {
                                Workbook = Path.GetFileName(inputPath),
                                Worksheet = sheet.Name,
                                ChartIndex = chartIdx,
                                ChartType = chart.Type.ToString(),
                                SeriesNames = seriesNames,
                                AxisTitles = new
                                {
                                    ValueAxis = valueAxisTitle,
                                    CategoryAxis = categoryAxisTitle
                                }
                            };

                            // Serialize to JSON (indented)
                            string json = JsonSerializer.Serialize(metadata, new JsonSerializerOptions { WriteIndented = true });

                            // Determine JSON file name
                            string jsonFileName = $"{Path.GetFileNameWithoutExtension(inputPath)}_Sheet{wsIndex + 1}_Chart{chartIdx + 1}.json";
                            string jsonPath = Path.Combine(outputFolder, jsonFileName);
                            File.WriteAllText(jsonPath, json);
                            Console.WriteLine($"Metadata for chart {chartIdx + 1} written to \"{jsonFileName}\".");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error processing chart {chartIdx + 1} on sheet \"{sheet.Name}\": {ex.Message}");
                        }
                    }
                }

                Console.WriteLine("Processing completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
