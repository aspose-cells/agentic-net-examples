// Title: Export Aspose.Cells Chart to High‑Resolution JPEG (C#) – 300 DPI for PowerPoint
// Description: Creates a workbook, adds a column chart, sets ImageOrPrintOptions (JPEG, 300 DPI, 90 % quality) and uses chart.ToImage to generate a high‑resolution JPEG suitable for PowerPoint slides.
// Keywords: Aspose.Cells | C# | chart export | high resolution JPEG | 300 DPI | ImageOrPrintOptions | ToImage | PowerPoint | Excel chart image
// Common Searches: Aspose.Cells export chart JPEG C# | high DPI chart image Aspose | 300 DPI Excel chart to JPEG | C# chart ToImage high quality | export Aspose chart for PowerPoint
// Developer Intent: Produce a 300 DPI JPEG image of an Excel chart for presentation use.
// Use Cases: Generate a column chart from worksheet data and export it as a 300 DPI JPEG for slide decks. | Create high‑quality chart images with adjustable JPEG quality to balance size and clarity. | Export chart images while keeping the original workbook editable for further processing.
// AI Prompts: Write C# code that loops through all charts in a workbook and saves each as a 300 DPI JPEG using Aspose.Cells. | Show how to export an Aspose.Cells chart as a PNG with transparent background and custom DPI. | Explain how to configure ImageOrPrintOptions for TIFF lossless export of a chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// Creates a workbook, adds a column chart, sets ImageOrPrintOptions (JPEG, 300 DPI, 90 % quality) and uses chart.ToImage to generate a high‑resolution JPEG suitable for PowerPoint slides.
class ExportChartHighResJpeg
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

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Configure high‑resolution JPEG options
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = ImageType.Jpeg;        // JPEG format
        options.HorizontalResolution = 300;       // 300 DPI horizontal
        options.VerticalResolution = 300;         // 300 DPI vertical
        options.Quality = 90;                     // JPEG quality (0‑100)

        // Export the chart as a high‑resolution JPEG image
        chart.ToImage("HighResChart.jpg", options);

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("ChartWorkbook.xlsx");
    }
}
