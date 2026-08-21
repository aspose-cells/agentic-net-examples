// Title: Export Aspose.Cells Chart to 300 DPI PNG and Save Next to Workbook (C#)
// Description: This C# example creates a workbook, adds sample data, builds a column chart, configures ImageOrPrintOptions for PNG with 300 dpi horizontal and vertical resolution, writes the chart to "ChartImage.png", and saves the workbook as "ChartWorkbook.xlsx" in the same directory.
// Keywords: Aspose.Cells chart export PNG | C# 300 dpi image | ImageOrPrintOptions DPI | save chart image with workbook | Aspose.Cells ToImage method | high resolution chart image .NET
// Common Searches: Aspose.Cells export chart PNG C# | set DPI for chart image Aspose.Cells | save chart image alongside workbook .NET | ImageOrPrintOptions 300 dpi example | convert Excel chart to high resolution PNG
// Developer Intent: Generate a PNG file of a worksheet chart at 300 dpi and keep the image file in the same folder as the Excel workbook.
// Use Cases: Produce a printable PDF that contains a crisp chart image while also delivering the original Excel file for further editing. | Create a web gallery of chart thumbnails, storing each PNG next to its source workbook for easy reference. | Automate batch extraction of all charts from a workbook as 300 dpi PNGs for documentation or presentation decks.
// AI Prompts: Write C# code that iterates through every chart in an Aspose.Cells workbook and saves each as a 300 dpi PNG in the workbook's directory. | Explain how ImageOrPrintOptions controls horizontal and vertical DPI when converting an Aspose.Cells chart to an image, with sample code. | Adapt the provided snippet to output a JPEG at 150 dpi instead of a PNG, while still saving the workbook in the same folder.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// This C# example creates a workbook, adds sample data, builds a column chart, configures ImageOrPrintOptions for PNG with 300 dpi horizontal and vertical resolution, writes the chart to "ChartImage.png", and saves the workbook as "ChartWorkbook.xlsx" in the same directory.
class ExportChart
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
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(15);
        worksheet.Cells["B4"].PutValue(7);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Configure image options: PNG format with 300 DPI resolution
        ImageOrPrintOptions imageOptions = new ImageOrPrintOptions();
        imageOptions.ImageType = ImageType.Png;
        imageOptions.HorizontalResolution = 300;
        imageOptions.VerticalResolution = 300;

        // Export the chart as a PNG image (saved alongside the workbook)
        string chartImagePath = "ChartImage.png";
        chart.ToImage(chartImagePath, imageOptions);

        // Save the workbook in the same folder
        string workbookPath = "ChartWorkbook.xlsx";
        workbook.Save(workbookPath);
    }
}
