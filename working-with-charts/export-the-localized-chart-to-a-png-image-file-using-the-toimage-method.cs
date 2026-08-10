// Title: Export Aspose.Cells Chart to PNG with C# Using the ToImage Method
// Description: This example creates a workbook, fills cells with category and sales data, adds a column chart based on range A1:B4, and uses the chart's ToImage method (ImageType.Png) to generate a PNG file named ChartImage.png. The workbook is then saved for reference.
// Keywords: Aspose.Cells | C# chart export | ToImage | PNG | Export chart as image | Aspose.Cells chart image | Aspose.Cells .NET | Chart to PNG | Save chart image | Aspose.Cells example
// Common Searches: Aspose.Cells export chart to PNG C# | How to use ToImage for chart in Aspose.Cells | Save Aspose.Cells chart as PNG file | C# Aspose.Cells chart image export example | Export chart image with Aspose.Cells .NET
// Developer Intent: Generate a PNG image from a chart in an Aspose.Cells workbook using C#.
// Use Cases: Embedding chart PNGs in web reports or dashboards. | Creating thumbnail previews of Excel charts for documentation. | Automating batch conversion of workbook charts to image files for publishing. | Providing chart images for email newsletters without sharing the full workbook.
// AI Prompts: Write C# code that loops through all charts in an Aspose.Cells workbook and saves each as a high‑resolution PNG using ToImage. | Explain how to set the image width, height, and DPI when exporting a chart with Aspose.Cells. | Show how to export a chart to JPEG instead of PNG with adjustable quality settings. | Provide a step‑by‑step guide to export a chart to PNG and then embed it into a PDF using Aspose.Pdf.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// This example creates a workbook, fills cells with category and sales data, adds a column chart based on range A1:B4, and uses the chart's ToImage method (ImageType.Png) to generate a PNG file named ChartImage.png. The workbook is then saved for reference.
class ExportChartToPng
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

        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["B2"].PutValue(1200);
        sheet.Cells["B3"].PutValue(800);
        sheet.Cells["B4"].PutValue(1500);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.SetChartDataRange("A1:B4", true);

        // Export the chart to a PNG image file using the ToImage method
        chart.ToImage("ChartImage.png", ImageType.Png);

        // Optionally save the workbook for reference
        workbook.Save("ChartWorkbook.xlsx");

        Console.WriteLine("Chart exported to ChartImage.png successfully.");
    }
}
