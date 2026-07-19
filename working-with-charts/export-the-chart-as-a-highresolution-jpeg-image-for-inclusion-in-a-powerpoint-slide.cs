// Title: Export Aspose.Cells Chart as High‑Resolution JPEG for PowerPoint (C#)
// Description: This C# example creates a workbook, builds a column chart from sample data, configures ImageOrPrintOptions with 300 DPI and 95 % quality, and saves the chart as a JPEG file ready for slide insertion.
// Keywords: Aspose.Cells | C# | chart export | JPEG | 300 DPI | high resolution image | ImageOrPrintOptions | PowerPoint | ToImage | column chart
// Common Searches: Aspose.Cells export chart JPEG | C# save chart as high DPI image | set JPEG quality Aspose.Cells | chart image for PowerPoint using .NET | increase chart resolution Aspose
// Developer Intent: Create a JPEG image of a worksheet chart with print‑quality resolution.
// Use Cases: Generate a slide‑ready chart image for presentations | Batch export multiple charts with uniform DPI for printing | Produce compact JPEGs that retain visual clarity for email reports
// AI Prompts: Write C# code to export an Aspose.Cells chart as a 600 DPI PNG with transparency. | Show how to loop through all charts in a workbook and save each as a high‑quality JPEG with custom filenames. | Explain how to balance JPEG compression and image size when exporting charts with ImageOrPrintOptions.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// This C# example creates a workbook, builds a column chart from sample data, configures ImageOrPrintOptions with 300 DPI and 95 % quality, and saves the chart as a JPEG file ready for slide insertion.
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
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Set high‑resolution JPEG options
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = ImageType.Jpeg;        // JPEG format
        options.HorizontalResolution = 300;       // 300 DPI horizontal
        options.VerticalResolution = 300;         // 300 DPI vertical
        options.Quality = 95;                     // JPEG quality (0‑100)

        // Export the chart as a high‑resolution JPEG image
        string outputImagePath = "HighResChart.jpg";
        chart.ToImage(outputImagePath, options);

        Console.WriteLine($"Chart exported successfully to '{outputImagePath}' with 300 DPI and quality 95.");
    }
}
