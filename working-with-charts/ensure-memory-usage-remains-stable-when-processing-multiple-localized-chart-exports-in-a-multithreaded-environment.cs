// Title: Parallel export of localized Excel charts to PNG with stable memory using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads each workbook with LoadOptions.MemorySetting set to MemoryPreference, applies a CultureInfo locale, and exports the first chart to a PNG file inside a Parallel.ForEach loop. | Enhance the parallel chart export method to accept a custom chart index and an output image format (PNG, JPEG) while guaranteeing all Aspose.Cells objects are disposed via using statements. | Add a CancellationToken parameter to ExportCharts so that ongoing chart exports can be cancelled gracefully without leaking memory.
// Common Searches: how to export Excel chart as PNG in a multithreaded C# application using Aspose.Cells | Aspose.Cells memory preference for processing many workbooks concurrently | set workbook locale for chart titles when exporting images with Aspose.Cells | prevent out-of-memory errors while exporting charts in parallel with Aspose.Cells | parallel.ForEach chart export example Aspose.Cells .NET
// Tags: multithreaded chart image generation Aspose.Cells | memory‑efficient workbook loading Aspose.Cells | localized chart title Aspose.Cells | chart to PNG conversion C# | using statement resource disposal Aspose.Cells

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example processes a collection of export requests in parallel, loading each workbook with a memory‑saving setting, applying the requested locale, retrieving or creating the first chart, localizing its title, binding it to a sample data range, and saving the chart as a PNG image, while using 'using' blocks to ensure native resources are released and memory usage stays stable.
    public class LocalizedChartExporter
    {
        // Represents a request to export a chart for a specific workbook and locale.
        public class ExportRequest
        {
            public string? WorkbookPath { get; set; }      // Path to the source Excel file
            public string? OutputImagePath { get; set; }   // Path where the chart image will be saved
            public string? Locale { get; set; }            // Locale identifier (e.g., "en-US", "fr-FR")
        }

        // Entry point for processing multiple export requests in parallel.
        public void ExportCharts(IEnumerable<ExportRequest> requests, int maxDegreeOfParallelism = 4)
        {
            var options = new ParallelOptions { MaxDegreeOfParallelism = maxDegreeOfParallelism };

            Parallel.ForEach(requests, options, request =>
            {
                try
                {
                    // Validate input parameters.
                    if (string.IsNullOrWhiteSpace(request?.WorkbookPath) ||
                        string.IsNullOrWhiteSpace(request?.OutputImagePath) ||
                        string.IsNullOrWhiteSpace(request?.Locale))
                    {
                        Console.WriteLine("Invalid request parameters; skipping.");
                        return;
                    }

                    // Ensure the source workbook exists.
                    if (!File.Exists(request.WorkbookPath))
                    {
                        Console.WriteLine($"Workbook not found: {request.WorkbookPath}");
                        return;
                    }

                    // Load the workbook with memory‑saving options.
                    var loadOptions = new LoadOptions(LoadFormat.Xlsx)
                    {
                        MemorySetting = MemorySetting.MemoryPreference
                    };

                    // Use 'using' to guarantee native resources are released.
                    using (var workbook = new Workbook(request.WorkbookPath, loadOptions))
                    {
                        // Apply the requested locale (affects number formats, etc.).
                        workbook.Settings.CultureInfo = new CultureInfo(request.Locale);

                        // Assume the first worksheet contains the data and chart.
                        if (workbook.Worksheets.Count == 0)
                        {
                            Console.WriteLine("No worksheets found in workbook.");
                            return;
                        }

                        var sheet = workbook.Worksheets[0];

                        // Retrieve an existing chart or create a new one.
                        Chart chart;
                        if (sheet.Charts.Count > 0)
                        {
                            chart = sheet.Charts[0];
                        }
                        else
                        {
                            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                            chart = sheet.Charts[chartIndex];
                        }

                        // Localize chart title.
                        chart.Title.Text = $"Sales Report ({request.Locale})";

                        // Bind chart series to a sample data range (A1:A10 in this example).
                        int firstRow = 0, lastRow = 9, firstColumn = 0;
                        chart.NSeries.Clear();

                        // Build a range reference string like "Sheet1!A1:A10".
                        string startCell = CellsHelper.CellIndexToName(firstRow, firstColumn);
                        string endCell = CellsHelper.CellIndexToName(lastRow, firstColumn);
                        string rangeRef = $"{sheet.Name}!{startCell}:{endCell}";

                        chart.NSeries.Add(rangeRef, true);
                        chart.NSeries[0].Name = sheet.Cells[firstRow, firstColumn].StringValue;

                        // Export the chart to a PNG image file (default format is PNG).
                        chart.ToImage(request.OutputImagePath);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing request for workbook '{request?.WorkbookPath}': {ex.Message}");
                }
            });
        }
    }

    // Simple program entry point to demonstrate usage.
    public static class Program
    {
        public static void Main()
        {
            var exporter = new LocalizedChartExporter();

            var requests = new List<LocalizedChartExporter.ExportRequest>
            {
                new LocalizedChartExporter.ExportRequest
                {
                    WorkbookPath = "SampleData.xlsx",
                    OutputImagePath = "Chart_en-US.png",
                    Locale = "en-US"
                },
                new LocalizedChartExporter.ExportRequest
                {
                    WorkbookPath = "SampleData.xlsx",
                    OutputImagePath = "Chart_fr-FR.png",
                    Locale = "fr-FR"
                }
            };

            exporter.ExportCharts(requests);
        }
    }
}
