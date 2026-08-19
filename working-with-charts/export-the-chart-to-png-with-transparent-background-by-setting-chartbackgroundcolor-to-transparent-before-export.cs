// Title: Export Aspose.Cells Chart to Transparent PNG in C# (.NET)
// Description: Creates a workbook, adds sample data and a column chart, sets ChartArea.BackgroundMode to Transparent, configures ImageOrPrintOptions for PNG with transparency, and saves the chart as a PNG file with a transparent background using Aspose.Cells for .NET.
// Keywords: Aspose.Cells chart export PNG | transparent chart background C# | ChartArea.BackgroundMode Transparent | ImageOrPrintOptions Transparent PNG | Aspose.Cells .NET chart image
// Common Searches: Aspose.Cells export chart as PNG with transparent background | C# make chart background transparent before saving | How to save Aspose.Cells chart to PNG with no background | Transparent PNG chart Aspose.Cells example
// Developer Intent: Generate a PNG image of a chart with a transparent background using Aspose.Cells for .NET.
// Use Cases: Embedding chart images on web pages where the page background should show through. | Creating overlay graphics for presentations or PDFs without a white box. | Producing theme‑aware chart assets for UI components that support transparency.
// AI Prompts: Provide a C# code sample that sets ChartArea.BackgroundMode to Transparent and exports the chart as a PNG using Aspose.Cells. | Explain how to enable transparent PNG output when saving a chart with ImageOrPrintOptions in Aspose.Cells. | Show step‑by‑step instructions to create a chart, make its background transparent, and save it as a PNG file in .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// Creates a workbook, adds sample data and a column chart, sets ChartArea.BackgroundMode to Transparent, configures ImageOrPrintOptions for PNG with transparency, and saves the chart as a PNG file with a transparent background using Aspose.Cells for .NET.
class ExportChartTransparent
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Fill sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["A4"].PutValue("Banana");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(80);
        sheet.Cells["B4"].PutValue(150);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Make the chart background transparent
        chart.ChartArea.BackgroundMode = BackgroundMode.Transparent;

        // Configure image options: PNG format with transparent background
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
        imgOptions.ImageType = ImageType.Png;   // Ensure PNG output
        imgOptions.Transparent = true;         // Enable transparent background

        // Export the chart to a PNG file with transparent background
        chart.ToImage("transparent_chart.png", imgOptions);
    }
}
