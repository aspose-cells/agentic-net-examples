// Title: How to measure Aspose.Cells chart rendering time after disabling data label text wrapping in C#
// AI Prompts: Generate C# code that creates a column chart with Aspose.Cells, disables data label text wrapping, renders the chart to a MemoryStream, and records the elapsed time using Stopwatch. | Show how to configure ImageOrPrintOptions for chart image export and output the rendering duration in milliseconds for performance testing. | Provide a complete example that saves the workbook, prints the chart rendering time, and handles any exceptions.
// Common Searches: aspnet measure chart rendering performance with Aspose.Cells | disable data label text wrap Aspose.Cells chart speed test | C# stopwatch chart image generation Aspose.Cells benchmark | how to time Aspose.Cells chart to image conversion | chart rendering latency after turning off data labels wrapping in .NET
// Tags: chart rendering performance Aspose.Cells C# | disable data label text wrap Aspose.Cells chart | measure chart image generation time Stopwatch | export Aspose.Cells column chart to image benchmark | Aspose.Cells workbook save with chart timing

using System;
using System.Diagnostics;
using System.IO;
using System.Drawing.Imaging;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// Creates a workbook, adds sample data, builds a column chart, disables data label text wrapping, renders the chart to an image while timing the operation with Stopwatch, outputs the elapsed milliseconds, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series
            chart.NSeries.Add("B2:B4", true);

            // Enable data labels for the series
            chart.NSeries[0].DataLabels.ShowValue = true;          // Show the value on each data label
            chart.NSeries[0].DataLabels.IsTextWrapped = false;    // Disable text wrapping for performance

            // Prepare image rendering options (default format is PNG)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();

            // Measure the time taken to render the chart to an image
            Stopwatch timer = new Stopwatch();
            timer.Start();

            using (MemoryStream ms = new MemoryStream())
            {
                chart.ToImage(ms, imgOptions);
                // Optionally, write the image to a file for verification
                // File.WriteAllBytes("ChartImage.png", ms.ToArray());
            }

            timer.Stop();
            Console.WriteLine($"Chart rendering time (ms): {timer.ElapsedMilliseconds}");

            // Save the workbook to a file
            string outputPath = "ChartRenderingTime.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
