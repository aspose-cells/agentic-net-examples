// Title: Export a column chart from Aspose.Cells to a PNG file and store it in a custom output folder using C#
// AI Prompts: Generate C# code that creates a workbook, adds a column chart with sample data, and uses Aspose.Cells' Chart.ToImage method to write the chart as a PNG to a specified directory. | Show how to programmatically ensure an output folder exists and then save an Aspose.Cells chart image (PNG) to that folder in a C# console application.
// Common Searches: C# Aspose.Cells how to save a chart as a PNG image file | export Excel column chart to PNG using Aspose.Cells library | Aspose.Cells Chart.ToImage example with output directory in C# | create and export chart image from workbook with Aspose.Cells C# | save Aspose.Cells chart to file path programmatically
// Tags: Aspose.Cells chart export PNG | C# Aspose.Cells column chart image generation | Chart.ToImage PNG output | ensure output directory Aspose.Cells C# | save Excel chart as image using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example creates a new workbook, fills cells A1:B4 with sample data, adds a column chart, defines its data range, creates an "output" folder if needed, and uses Chart.ToImage to export the chart as a PNG file named ChartImage.png.
class ExportChartToPng
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
        worksheet.Cells["B1"].PutValue("Sales");
        worksheet.Cells["B2"].PutValue(1200);
        worksheet.Cells["B3"].PutValue(800);
        worksheet.Cells["B4"].PutValue(1500);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.SetChartDataRange("A1:B4", true);

        // Ensure the output directory exists
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Export the chart as a PNG image
        string imagePath = Path.Combine(outputDir, "ChartImage.png");
        chart.ToImage(imagePath, ImageType.Png);

        Console.WriteLine($"Chart exported successfully to: {imagePath}");
    }
}
