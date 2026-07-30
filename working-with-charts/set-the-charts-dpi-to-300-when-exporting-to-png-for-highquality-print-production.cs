// Title: Export Aspose.Cells Chart to 300 DPI PNG in C# for Print‑Ready Output
// Description: Creates a workbook, adds a column chart, configures ImageOrPrintOptions with ImageType.Png and 300 dpi horizontal/vertical resolution, then uses chart.ToImage to generate a print‑quality PNG file.
// Keywords: Aspose.Cells | C# chart export | 300 DPI PNG | ImageOrPrintOptions | HorizontalResolution | VerticalResolution | print quality chart | Excel chart to image | high DPI export | .NET Aspose.Cells
// Common Searches: Aspose.Cells export chart PNG 300 DPI | Set chart image resolution in Aspose.Cells C# | Print‑ready chart image from Excel using Aspose.Cells | ImageOrPrintOptions DPI settings example | C# export Excel chart as high‑resolution PNG
// Developer Intent: Generate a PNG image of a worksheet chart at 300 dpi so it can be used in print‑oriented documents.
// Use Cases: Produce a column chart image for a marketing brochure directly from a .NET application. | Automate high‑resolution chart generation in a nightly reporting pipeline. | Create print‑ready graphics for financial statements without manual resizing.
// AI Prompts: Write C# code with Aspose.Cells to export a chart as a 600 DPI PNG. | Show how to export a chart to JPEG at 300 DPI and embed it in a PDF using Aspose.Pdf. | Explain how to loop through all charts in a workbook and save each as a 300 DPI PNG.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// Creates a workbook, adds a column chart, configures ImageOrPrintOptions with ImageType.Png and 300 dpi horizontal/vertical resolution, then uses chart.ToImage to generate a print‑quality PNG file.
class ExportChartHighDPI
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Product A");
        sheet.Cells["A3"].PutValue("Product B");
        sheet.Cells["A4"].PutValue("Product C");

        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(80);
        sheet.Cells["B4"].PutValue(150);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Set image options: PNG format with 300 DPI resolution
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = ImageType.Png;
        options.HorizontalResolution = 300;
        options.VerticalResolution = 300;

        // Export the chart to a PNG file using the high‑DPI settings
        chart.ToImage("Chart_300dpi.png", options);
    }
}
