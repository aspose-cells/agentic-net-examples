// Title: Export Aspose.Cells Chart to 300 DPI PNG in C#
// Description: Creates a workbook, adds sample data, builds a column chart, configures ImageOrPrintOptions for 300 DPI PNG, and saves the chart as a high‑definition image file.
// Keywords: Aspose.Cells | C# chart export | PNG 300 DPI | ImageOrPrintOptions | high resolution chart image | set chart DPI | Aspose.Cells .NET | chart to image | 300 DPI PNG | export chart as PNG
// Common Searches: Aspose.Cells export chart PNG 300 DPI C# | set chart image resolution Aspose.Cells | high resolution chart image .NET | ImageOrPrintOptions DPI setting | save Aspose.Cells chart as high quality PNG
// Developer Intent: Generate a 300 DPI PNG image of a workbook chart.
// Use Cases: Print‑ready charts for marketing brochures | High‑resolution visuals for PowerPoint or PDF reports | Sharp thumbnails for analytics dashboards | Compliance‑level graphics for regulatory documents
// AI Prompts: How to change the DPI to 600 and output JPEG instead of PNG using Aspose.Cells? | Show code to batch export all charts in a workbook to 300 DPI PNG files. | Explain how to retrieve the actual DPI of a saved chart image with Aspose.Cells. | Provide a method to embed the exported PNG into a PDF with matching DPI.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartDpiExample
{
    // Creates a workbook, adds sample data, builds a column chart, configures ImageOrPrintOptions for 300 DPI PNG, and saves the chart as a high‑definition image file.
    class Program
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

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure image options for high‑definition PNG (300 DPI)
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png,          // Ensure PNG format
                HorizontalResolution = 300,        // 300 DPI horizontally
                VerticalResolution = 300           // 300 DPI vertically
            };

            // Save the chart as a PNG image with the specified DPI
            chart.ToImage("chart_300dpi.png", options);

            Console.WriteLine("Chart saved as PNG with 300 DPI resolution.");
        }
    }
}
