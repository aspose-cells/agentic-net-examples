// Title: Export Aspose.Cells Chart to 300 DPI PNG in C# for Print‑Ready Output
// Description: This example creates a workbook, adds sample sales data, builds a column chart, and uses ImageOrPrintOptions with ImageType.Png, HorizontalResolution = 300 and VerticalResolution = 300 to generate a print‑quality PNG image of the chart. The workbook can also be saved for further processing.
// Keywords: Aspose.Cells | C# chart export | 300 DPI PNG | ImageOrPrintOptions | high resolution chart image | print ready graphics | chart ToImage method | horizontalresolution | verticalresolution
// Common Searches: Aspose.Cells export chart 300 DPI PNG C# | set chart image resolution Aspose.Cells | how to increase DPI of chart PNG in .NET | ImageOrPrintOptions HorizontalResolution VerticalResolution usage | print quality chart export Aspose.Cells
// Developer Intent: Generate a PNG image of a workbook chart at 300 DPI so it meets print‑production quality standards.
// Use Cases: Produce crisp sales charts for marketing brochures and flyers. | Create high‑resolution graphics for inclusion in PDF or InDesign layouts. | Export chart assets for product catalogs, posters, or large‑format prints.
// AI Prompts: Show how to export an Aspose.Cells chart to a 600 DPI PNG using ImageOrPrintOptions. | Explain the impact of HorizontalResolution and VerticalResolution on file size and visual quality. | Provide code to batch‑export all charts in a workbook, assigning a custom DPI to each.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;   // For ImageType enum

// This example creates a workbook, adds sample sales data, builds a column chart, and uses ImageOrPrintOptions with ImageType.Png, HorizontalResolution = 300 and VerticalResolution = 300 to generate a print‑quality PNG image of the chart. The workbook can also be saved for further processing.
class ExportChartHighDpi
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Product A");
            sheet.Cells["A3"].PutValue("Product B");
            sheet.Cells["A4"].PutValue("Product C");

            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["B4"].PutValue(150);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure image options for high‑resolution PNG (300 DPI)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png,
                HorizontalResolution = 300, // 300 DPI horizontally
                VerticalResolution = 300    // 300 DPI vertically
            };

            // Export the chart to PNG using the high‑DPI settings
            chart.ToImage("Chart_300dpi.png", imgOptions);

            // Optionally, save the workbook if further processing is needed
            workbook.Save("Workbook_With_Chart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
