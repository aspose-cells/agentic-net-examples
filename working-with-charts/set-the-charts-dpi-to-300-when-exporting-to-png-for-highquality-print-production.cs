// Title: Export a chart to high‑resolution PNG (300 DPI) with Aspose.Cells for .NET
// Description: Creates a workbook, adds a column chart from sample data, configures ImageOrPrintOptions for PNG with 300 DPI horizontal and vertical resolution, and uses Chart.ToImage to generate a print‑ready PNG file.
// Keywords: Aspose.Cells chart DPI | C# export chart PNG 300 DPI | ImageOrPrintOptions resolution | high resolution chart image Aspose.Cells | chart ToImage .NET | print ready chart PNG | Aspose.Cells export PNG
// Common Searches: set DPI when exporting chart to PNG Aspose.Cells C# | Aspose.Cells high resolution chart image example | Chart.ToImage 300 DPI PNG .NET | how to change chart export resolution Aspose.Cells | print quality chart PNG Aspose.Cells
// Developer Intent: Configure a chart’s DPI to 300 and export it as a PNG image.
// Use Cases: Produce print‑ready chart graphics for reports and brochures. | Create high‑resolution chart assets for marketing or publishing. | Batch‑export multiple workbook charts with a uniform 300 DPI setting.
// AI Prompts: Generate C# code that exports every chart in a workbook to 300 DPI PNG files using Aspose.Cells. | Explain how to adjust DPI for chart images when saving to PNG, JPEG, or TIFF with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// Creates a workbook, adds a column chart from sample data, configures ImageOrPrintOptions for PNG with 300 DPI horizontal and vertical resolution, and uses Chart.ToImage to generate a print‑ready PNG file.
class SetChartDpiExample
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

        // Set image options: PNG format with 300 DPI resolution
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = ImageType.Png;
        options.HorizontalResolution = 300; // 300 DPI horizontally
        options.VerticalResolution = 300;   // 300 DPI vertically

        // Export the chart to a PNG file using the specified DPI
        chart.ToImage("Chart_300dpi.png", options);

        Console.WriteLine("Chart exported to PNG with 300 DPI.");
    }
}
