// Title: Export a column chart from Aspose.Cells to a 300 DPI JPEG for PowerPoint using C#
// AI Prompts: Write C# code that creates a column chart with Aspose.Cells and saves it as a 300 DPI JPEG image suitable for PowerPoint slides. | Show how to configure ImageOrPrintOptions for maximum JPEG quality and specific DPI when exporting an Aspose.Cells chart.
// Common Searches: c# export aspose.cells chart as 300 dpi jpeg for powerpoint | how to set jpeg quality when saving an aspose.cells chart | imageorprintoptions high resolution jpeg example aspose.cells | save column chart as high resolution jpeg using aspose.cells c# | export chart image for slide deck aspose.cells c#
// Tags: Aspose.Cells chart JPEG export | high‑resolution chart image C# | ImageOrPrintOptions 300 DPI | column chart to JPEG Aspose.Cells | chart export for PowerPoint slide

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds sample data, builds a column chart, configures ImageOrPrintOptions for 300 DPI and maximum JPEG quality, and exports the chart as a high‑resolution JPEG file ready for inclusion in a PowerPoint presentation.
class ExportChartHighResJpeg
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["A3"].PutValue("Orange");
        worksheet.Cells["A4"].PutValue("Banana");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(120);
        worksheet.Cells["B3"].PutValue(80);
        worksheet.Cells["B4"].PutValue(150);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Set high‑resolution JPEG options
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = ImageType.Jpeg;        // Output format
        options.HorizontalResolution = 300;       // 300 DPI horizontal
        options.VerticalResolution = 300;         // 300 DPI vertical
        options.Quality = 100;                    // Maximum JPEG quality

        // Export the chart as a high‑resolution JPEG image
        string outputPath = "HighResolutionChart.jpg";
        chart.ToImage(outputPath, options);

        Console.WriteLine($"Chart successfully exported to: {outputPath}");
    }
}
