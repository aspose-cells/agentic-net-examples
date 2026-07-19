// Title: Export an Aspose.Cells Chart to PNG with C# Using Chart.ToImage
// Description: Demonstrates how to create a workbook, add sample data, generate a column chart, and export the chart as a PNG image using the Chart.ToImage method (ImageType.Png) in Aspose.Cells for .NET. The workbook can also be saved for reference.
// Keywords: Aspose.Cells export chart PNG | Chart.ToImage C# | save Excel chart as image | Aspose.Cells chart image generation | C# export chart to PNG
// Common Searches: How to export an Aspose.Cells chart to PNG in C# | Aspose.Cells Chart.ToImage example | Save Excel chart as PNG using Aspose.Cells | Export chart image with Aspose.Cells .NET
// Developer Intent: Generate a PNG file from a chart created in an Aspose.Cells workbook using C#.
// Use Cases: Create PNG snapshots of sales charts for inclusion in PDF reports. | Produce localized chart thumbnails for web dashboards that respect regional formatting. | Automate chart image generation for email notifications or document workflows.
// AI Prompts: Show C# code that creates a line chart with Aspose.Cells and saves it as a JPEG using Chart.ToImage. | Provide an example that exports multiple charts from a workbook to separate PNG files with Aspose.Cells. | Explain how to control image resolution and dimensions when exporting a chart with Chart.ToImage.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace ExportChartToPng
{
    // Demonstrates how to create a workbook, add sample data, generate a column chart, and export the chart as a PNG image using the Chart.ToImage method (ImageType.Png) in Aspose.Cells for .NET. The workbook can also be saved for reference.
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

            worksheet.Cells["B1"].PutValue("Sales");
            worksheet.Cells["B2"].PutValue(1200);
            worksheet.Cells["B3"].PutValue(800);
            worksheet.Cells["B4"].PutValue(1500);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.SetChartDataRange("A1:B4", true);

            // Export the chart to a PNG image file
            string imagePath = "ChartImage.png";
            chart.ToImage(imagePath, ImageType.Png);

            Console.WriteLine($"Chart exported successfully to {imagePath}");

            // Optionally save the workbook for reference
            workbook.Save("ChartWorkbook.xlsx");
        }
    }
}
