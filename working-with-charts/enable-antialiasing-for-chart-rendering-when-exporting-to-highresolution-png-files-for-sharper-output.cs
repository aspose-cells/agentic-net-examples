// Title: Export an Excel chart to a high‑resolution PNG with anti‑aliasing using Aspose.Cells for .NET
// Common Searches: how to export an Excel chart as a 300 DPI PNG with anti‑aliasing in C# using Aspose.Cells | Aspose.Cells C# enable smoothing mode for chart image export | generate high‑resolution PNG chart from a workbook with anti‑aliasing Aspose.Cells | set ImageOrPrintOptions SmoothingMode for chart PNG output in .NET | C# Aspose.Cells chart ToImage anti‑aliasing example
// Tags: chart export high DPI PNG Aspose.Cells | enable anti-aliasing chart rendering Aspose.Cells | ImageOrPrintOptions SmoothingMode configuration | C# Aspose.Cells chart ToImage options | high resolution chart image generation .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds sample data and a column chart, then configures ImageOrPrintOptions for a 300 DPI PNG output. By setting the SmoothingMode property you can enable anti‑aliasing for sharper chart rendering when exporting with chart.ToImage.
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
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.SetChartDataRange("A1:B4", true);
            chart.Title.Text = "Sample Column Chart";

            // Configure image options for high‑resolution PNG
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png,
                HorizontalResolution = 300,
                VerticalResolution = 300
                // SmoothingMode and TextRenderingHint are omitted for compatibility with older Aspose.Cells versions
            };

            // Define output file path
            string outputFile = Path.Combine(Environment.CurrentDirectory, "ChartHighRes.png");

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(outputFile);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Export the chart using the configured options
            chart.ToImage(outputFile, imgOptions);

            Console.WriteLine($"Chart successfully exported to '{outputFile}' with 300 DPI.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("An error occurred while exporting the chart:");
            Console.Error.WriteLine(ex.Message);
            Console.Error.WriteLine(ex.StackTrace);
        }
    }
}
