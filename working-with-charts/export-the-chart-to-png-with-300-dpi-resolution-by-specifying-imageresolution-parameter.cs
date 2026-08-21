// Title: Export Aspose.Cells Chart to High‑Resolution PNG (300 DPI) Using ImageOrPrintOptions in C#
// Description: Demonstrates how to create a workbook, add a column chart, set ImageOrPrintOptions.HorizontalResolution and VerticalResolution to 300 DPI, and save the chart as a 300‑DPI PNG image with Aspose.Cells for .NET.
// Keywords: Aspose.Cells chart export PNG | 300 DPI image Aspose.Cells | ImageOrPrintOptions resolution | C# chart to image | high resolution Excel chart | chart.ToImage DPI | Aspose.Cells .NET example
// Common Searches: export Aspose.Cells chart as PNG 300 DPI | set chart image resolution in Aspose.Cells C# | ImageOrPrintOptions HorizontalResolution VerticalResolution | Aspose.Cells high‑resolution chart image | C# save Excel chart to PNG with custom DPI
// Developer Intent: Generate a PNG file of a worksheet chart at 300 DPI using Aspose.Cells for .NET.
// Use Cases: Produce print‑ready chart graphics for reports and brochures. | Create high‑quality PNG assets for dashboards and marketing collateral. | Automate batch conversion of multiple Excel charts to consistent 300 DPI images.
// AI Prompts: Write C# code that iterates through all charts in a workbook and exports each to a 300 DPI PNG using Aspose.Cells. | Show how to export a chart to JPEG with a custom DPI and transparent background via ImageOrPrintOptions. | Explain how to maintain aspect ratio while changing DPI when saving an Aspose.Cells chart as an image.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// Demonstrates how to create a workbook, add a column chart, set ImageOrPrintOptions.HorizontalResolution and VerticalResolution to 300 DPI, and save the chart as a 300‑DPI PNG image with Aspose.Cells for .NET.
class ExportChartPng
{
    static void Main()
    {
        try
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
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";      // Categories

            // Configure image options with 300 DPI resolution (default format is PNG)
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                HorizontalResolution = 300, // 300 DPI horizontal
                VerticalResolution = 300    // 300 DPI vertical
            };

            // Export the chart to a PNG file using the specified DPI
            string outputPath = "Chart_300dpi.png";
            chart.ToImage(outputPath, options);

            Console.WriteLine($"Chart exported to {outputPath} with 300 DPI resolution.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
