// Title: Export an Aspose.Cells chart to JPEG in C# using the ToImage method
// Description: Creates a workbook, fills cells A1:B4 with sample data, adds a column chart, and saves the chart as a JPEG file (ChartImage.jpeg) by calling Chart.ToImage with ImageType.Jpeg.
// Keywords: Aspose.Cells chart export | C# ToImage JPEG | save Aspose chart as image | export Excel chart to JPEG | .NET chart to image | Aspose.Cells image conversion | chart ToImage example | Aspose.Cells C# tutorial | Excel chart image generation
// Common Searches: Aspose.Cells export chart to JPEG C# | How to save a chart as JPEG using Aspose.Cells | ToImage method example for chart export | C# code to convert Excel chart to JPEG | Aspose.Cells chart image output
// Developer Intent: Generate a JPEG file from a chart created with Aspose.Cells.
// Use Cases: Embed sales charts as JPEG images in PDF reports. | Attach chart snapshots to automated email alerts. | Create thumbnail images for web dashboards that display Excel‑derived charts.
// AI Prompts: Show how to export the same chart as PNG instead of JPEG using ToImage. | Add a title and legend to the chart before saving it as a JPEG. | Write code that iterates over all charts in a worksheet and saves each as a separate JPEG file.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook, fills cells A1:B4 with sample data, adds a column chart, and saves the chart as a JPEG file (ChartImage.jpeg) by calling Chart.ToImage with ImageType.Jpeg.
class ExportChartToJpeg
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
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

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.SetChartDataRange("A1:B4", true);

        // Export the chart to a JPEG image file
        string imagePath = "ChartImage.jpeg";
        chart.ToImage(imagePath, ImageType.Jpeg);

        Console.WriteLine($"Chart exported successfully to {imagePath}");
    }
}
