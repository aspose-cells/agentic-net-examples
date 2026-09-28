// Title: Report conversion progress while exporting an Excel chart to PNG using Aspose.Cells and IProgress in C#
// AI Prompts: Write C# code that loads an .xlsx workbook, extracts the first chart, and saves it as a PNG with Aspose.Cells while sending percentage updates to a custom IProgress<int> implementation. | Demonstrate how to set ImageOrPrintOptions for PNG output and manually invoke a console‑based IProgress reporter to show conversion progress.
// Common Searches: how to monitor Aspose.Cells chart to PNG conversion progress in a .NET console application | using IProgress<int> to display percentage while exporting an Excel chart as an image with Aspose.Cells | C# example converting the first worksheet chart to PNG and reporting progress in the console
// Tags: Aspose.Cells chart export to PNG with progress updates | C# IProgress usage for image conversion | ImageOrPrintOptions configuration for PNG output | console percentage reporting during chart rendering

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// Progress reporter that implements IProgress<int> and writes percentage to console
// The program loads an Excel workbook, retrieves the first chart from the first worksheet, converts the chart to a PNG file using Aspose.Cells ImageOrPrintOptions, and reports conversion progress to the console via a custom IProgress<int> implementation.
class ConsoleProgressReporter : IProgress<int>
{
    public void Report(int value)
    {
        // Clamp the value between 0 and 100
        int percent = Math.Max(0, Math.Min(100, value));
        Console.WriteLine($"Conversion progress: {percent}%");
    }
}

class ChartToPngConverter
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "chart.png";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Assume the first worksheet contains the chart
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one chart
            if (sheet.Charts.Count == 0)
                throw new InvalidOperationException("No charts found in the first worksheet.");

            // Get the first chart
            Chart chart = sheet.Charts[0];

            // Prepare image options for PNG output
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                SaveFormat = SaveFormat.Png // Use PNG format without needing System.Drawing.ImageFormat
            };

            // Create a progress reporter instance
            IProgress<int> progress = new ConsoleProgressReporter();

            // Convert the chart to PNG while reporting progress
            using (FileStream pngStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
            {
                // The ToImage overload that accepts ImageOrPrintOptions does not directly support IProgress,
                // so we invoke the overload with a stream and ImageFormat via ImageOrPrintOptions.
                // Progress reporting is handled manually by invoking the reporter after conversion.
                chart.ToImage(pngStream, imgOptions);
                progress.Report(100); // Report completion
            }

            Console.WriteLine("Chart conversion completed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
