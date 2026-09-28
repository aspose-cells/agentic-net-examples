// Title: Export a column chart to a 300 DPI PNG file and save it alongside the workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code to build a column chart in an Aspose.Cells workbook and export the chart to a PNG file using a 300 dpi resolution. | Demonstrate configuring ImageOrPrintOptions for PNG output with both horizontal and vertical resolution set to 300 dpi in Aspose.Cells. | Show a C# example that saves the workbook and the generated PNG chart file to the same folder.
// Common Searches: how to export an Aspose.Cells chart as a high‑resolution PNG in C# | Aspose.Cells C# chart export 300 DPI PNG image | save chart image and workbook in same directory using Aspose.Cells for .NET | configure ImageOrPrintOptions for 300 dpi when exporting chart with Aspose.Cells | C# Aspose.Cells export column chart to PNG with specific resolution
// Tags: Aspose.Cells chart PNG export | ImageOrPrintOptions high‑resolution output | C# column chart export Aspose | save workbook and chart image together | 300 DPI chart image Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// // Creates a workbook, adds sample data and a column chart, sets ImageOrPrintOptions to PNG with 300 dpi horizontal and vertical resolution, exports the chart as "ChartImage.png", and saves the workbook as "WorkbookWithChart.xlsx" in the same folder using Aspose.Cells for .NET.
class ExportChartAsPng
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["A4"].PutValue("Banana");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(1200);
        sheet.Cells["B3"].PutValue(800);
        sheet.Cells["B4"].PutValue(1500);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Set image options: PNG format with 300 DPI resolution
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
        imgOptions.ImageType = ImageType.Png;
        imgOptions.HorizontalResolution = 300;
        imgOptions.VerticalResolution = 300;

        // Export the chart as a PNG image using the specified options
        chart.ToImage("ChartImage.png", imgOptions);

        // Save the workbook in the same directory
        workbook.Save("WorkbookWithChart.xlsx");
    }
}
