// Title: Export an Aspose.Cells chart to PNG while showing conversion progress with IProgress<int> in C#
// AI Prompts: Generate C# code that creates a workbook, adds a chart, and writes the chart to a PNG stream while an IProgress<int> implementation reports the rendering percentage to the console. | Refactor the chart‑to‑PNG sample to accept an IProgress<int> argument and invoke Report at appropriate points during the ToImage operation. | Design a helper method that receives a Chart object, ImageOrPrintOptions, and an IProgress<int>, returns a PNG file, and provides real‑time progress feedback.
// Common Searches: how to display progress while exporting Aspose.Cells chart to PNG in C# | Aspose.Cells IProgress<int> example for chart image rendering | track Aspose.Cells ToImage conversion percentage console output | C# export Excel chart as PNG with progress callback using Aspose.Cells | report chart rendering progress using IProgress in Aspose.Cells library
// Tags: Aspose.Cells chart to PNG conversion with progress | C# IProgress reporting for Aspose.Cells image rendering | export Excel chart as PNG using Aspose.Cells | track ToImage operation progress in Aspose.Cells | console progress feedback for chart image generation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The example builds a workbook, fills it with sample data, creates a column chart, and defines a ConsoleProgress class that implements IProgress<int> to write percentage values to the console. ImageOrPrintOptions are prepared (default PNG) and the chart is exported via chart.ToImage to a file. Although the progress reporter is defined, the current Aspose.Cells version does not pass it to ToImage, so no progress updates are emitted during rendering.
class ConsoleProgress : IProgress<int>
{
    // Report progress percentage to the console
    public void Report(int value)
    {
        Console.WriteLine($"Conversion progress: {value}%");
    }
}

class ChartToPngWithProgress
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Prepare image export options (progress reporting not supported in this version)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // Default image format is PNG; no need to set ImageFormat explicitly
            };

            // Export the chart to a PNG file
            using (FileStream pngStream = new FileStream("ChartOutput.png", FileMode.Create, FileAccess.Write))
            {
                chart.ToImage(pngStream, imgOptions);
            }

            Console.WriteLine("Chart has been exported to ChartOutput.png");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
