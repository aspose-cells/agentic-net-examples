// Title: Export Aspose.Cells Chart to PNG at 300 DPI (C#)
// Description: This C# sample shows how to build a workbook, add data, create a column chart, and use ImageOrPrintOptions to render the chart as a PNG image with both horizontal and vertical resolution set to 300 dpi. The chart image is saved separately while the workbook can also be persisted.
// Keywords: Aspose.Cells | C# chart export | PNG 300 DPI | ImageOrPrintOptions | high resolution chart image | chart.ToImage | Aspose.Cells rendering | set DPI | export chart as PNG | Aspose.Cells .NET
// Common Searches: Aspose.Cells export chart PNG 300 dpi | How to set image resolution for chart export in C# | ImageOrPrintOptions horizontalresolution verticalresolution example | Generate high‑resolution chart image with Aspose.Cells | C# chart.ToImage DPI setting
// Developer Intent: Create a 300‑dpi PNG file from an Aspose.Cells chart.
// Use Cases: Produce print‑ready chart graphics for marketing brochures. | Supply high‑quality PNG charts for dashboards that require exact pixel dimensions. | Automate batch export of workbook charts at a specific DPI for documentation. | Integrate chart images into PDF reports where 300 dpi ensures clarity.
// AI Prompts: Generate C# code to export an Aspose.Cells chart to a 600 dpi TIFF using ImageOrPrintOptions. | Write a script that iterates through all charts in a workbook and saves each as a PNG with a custom DPI value. | Explain how to retrieve the file path returned by chart.ToImage and handle possible exceptions.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// This C# sample shows how to build a workbook, add data, create a column chart, and use ImageOrPrintOptions to render the chart as a PNG image with both horizontal and vertical resolution set to 300 dpi. The chart image is saved separately while the workbook can also be persisted.
class ExportChartPng300Dpi
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["A4"].PutValue("Banana");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(80);
        sheet.Cells["B4"].PutValue(150);

        // Insert a column chart and set its data source
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Configure image options: PNG format with 300 DPI resolution
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = ImageType.Png;
        options.HorizontalResolution = 300;
        options.VerticalResolution = 300;

        // Export the chart to a PNG file using the specified resolution
        chart.ToImage("Chart_300dpi.png", options);

        // Save the workbook (optional, demonstrates normal lifecycle)
        workbook.Save("ChartWorkbook.xlsx");
    }
}
