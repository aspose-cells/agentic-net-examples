// Title: Monitor Aspose.Cells chart‑to‑PNG conversion with IProgress in C#
// Description: Creates a workbook, adds a column chart, and uses ImageOrPrintOptions with a custom IPageSavingCallback to report the rendering percentage via an IProgress<int> instance that writes to the console while saving the chart as a PNG file.
// Keywords: Aspose.Cells | chart export PNG | IProgress<int> | IPageSavingCallback | C# | .NET | image conversion progress | console logging | ImageOrPrintOptions | chart rendering feedback
// Common Searches: Aspose.Cells track chart export progress C# | IProgress with ImageOrPrintOptions example | How to show percentage while saving chart as PNG | Implement IPageSavingCallback for image conversion | Console progress report for Aspose.Cells chart rendering
// Developer Intent: Export a worksheet chart to a PNG image and display the conversion percentage in the console during the process.
// Use Cases: Command‑line tools that need real‑time feedback when converting large or multi‑page charts. | Batch jobs that process many charts and require per‑file progress reporting for monitoring. | Desktop applications that want to inform users about chart image generation to improve perceived performance.
// AI Prompts: Write a C# snippet that saves an Aspose.Cells chart as JPEG and reports progress with IProgress<int>. | Show how to modify the ChartProgressCallback to write progress updates to a log file instead of the console. | Provide an example that adds CancellationToken support to the chart‑to‑image conversion while still using IProgress for status updates.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// Creates a workbook, adds a column chart, and uses ImageOrPrintOptions with a custom IPageSavingCallback to report the rendering percentage via an IProgress<int> instance that writes to the console while saving the chart as a PNG file.
class Program
{
    static void Main()
    {
        // Create a workbook and add sample data for the chart
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["A3"].PutValue("Banana");
        worksheet.Cells["A4"].PutValue("Cherry");

        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(30);
        worksheet.Cells["B3"].PutValue(45);
        worksheet.Cells["B4"].PutValue(25);

        // Add a column chart and bind it to the data range
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Set up a progress reporter that writes percentage to the console
        IProgress<int> progress = new Progress<int>(p => Console.WriteLine($"Conversion progress: {p}%"));

        // Configure image options and attach the custom page‑saving callback
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Png,
            PageSavingCallback = new ChartProgressCallback(progress)
        };

        // Convert the chart to a PNG file while reporting progress
        chart.ToImage("chart.png", options);
    }

    // Custom callback that reports conversion progress via IProgress<int>
    private class ChartProgressCallback : IPageSavingCallback
    {
        private readonly IProgress<int> _progress;

        public ChartProgressCallback(IProgress<int> progress) => _progress = progress;

        public void PageStartSaving(PageStartSavingArgs args)
        {
            // Calculate percentage based on current page index and total page count
            int percent = (int)((args.PageIndex + 1) * 100.0 / args.PageCount);
            _progress.Report(percent);
        }

        public void PageEndSaving(PageEndSavingArgs args)
        {
            // Ensure the final report reaches 100%
            _progress.Report(100);
        }
    }
}
