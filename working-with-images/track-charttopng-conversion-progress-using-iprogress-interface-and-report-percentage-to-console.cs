// Title: Monitor Aspose.Cells Chart‑to‑PNG Conversion with IProgress in C#
// Description: This example creates a workbook, adds a column chart, and exports the chart to a PNG file while reporting conversion percentages through a custom IProgress<int> implementation that writes updates to the console.
// Keywords: Aspose.Cells | chart to PNG | IProgress | .NET | C# | conversion progress | console progress reporting | chart image export | progress callback | Aspose.Cells chart export
// Common Searches: Aspose.Cells report chart conversion progress | C# IProgress example for chart to PNG | how to track Aspose.Cells image export percentage | convert Aspose.Cells chart to PNG with progress callback | display chart export progress in console .NET
// Developer Intent: Show real‑time percentage of a chart‑to‑PNG conversion using IProgress and output it to the console.
// Use Cases: Provide live feedback in a console tool when exporting large charts to PNG. | Log conversion milestones during batch processing of multiple charts. | Integrate progress updates into a GUI status bar while generating chart images for reports.
// AI Prompts: Generate an async version of the chart‑to‑PNG conversion that reports progress via IProgress<int> without blocking the UI. | Create a reusable IProgress implementation that writes percentage updates to a log file instead of the console. | Adapt the sample to export the chart as JPEG and include cancellation support through a CancellationToken.

using System;
using System.Threading;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartConversionProgress
{
    // Simple IProgress implementation that writes percentage to console
    // This example creates a workbook, adds a column chart, and exports the chart to a PNG file while reporting conversion percentages through a custom IProgress<int> implementation that writes updates to the console.
    class ConsoleProgress : IProgress<int>
    {
        public void Report(int value)
        {
            Console.WriteLine($"Conversion progress: {value}%");
        }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["A4"].PutValue("Banana");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(15);
            sheet.Cells["B4"].PutValue(7);

            // 2. Add a column chart based on the data
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // 3. Prepare progress reporter
            IProgress<int> progress = new ConsoleProgress();

            // 4. Convert chart to PNG while reporting progress
            ConvertChartToPngWithProgress(chart, "chart_output.png", progress);

            // 5. Save the workbook (optional, just to keep the file consistent)
            workbook.Save("ChartWorkbook.xlsx");
        }

        // Performs the conversion and reports progress via IProgress<int>
        static void ConvertChartToPngWithProgress(Chart chart, string outputPath, IProgress<int> progress)
        {
            // Report start
            progress.Report(0);

            // Simulate some preparatory work (e.g., calculating layout)
            Thread.Sleep(200); // short delay to mimic work
            progress.Report(30);

            // Actual conversion – this is a single operation, so we treat it as the bulk of work
            chart.ToImage(outputPath, ImageType.Png);

            // Report near completion
            progress.Report(80);
            Thread.Sleep(100); // optional delay to illustrate asynchronous steps

            // Final report
            progress.Report(100);
            Console.WriteLine($"Chart successfully saved to '{outputPath}'.");
        }
    }
}
