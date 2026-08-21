// Title: Export Aspose.Cells Chart to a High‑Resolution PNG (300 DPI) using C#
// Description: Creates a workbook, adds sample data, builds a column chart, sets ImageOrPrintOptions to 300 DPI, and saves the chart as a high‑resolution PNG. The workbook can also be saved for reference.
// Keywords: Aspose.Cells chart export PNG | high DPI chart image C# | ImageOrPrintOptions 300 DPI | Chart.ToImage high resolution | Aspose.Cells rendering options
// Common Searches: Aspose.Cells export chart to 300 DPI PNG | C# high‑resolution chart image Aspose.Cells | Set DPI for chart PNG in Aspose.Cells | How to render chart as high quality PNG with Aspose
// Developer Intent: Generate a 300 DPI PNG of a workbook chart to ensure crisp text and graphics.
// Use Cases: Print‑ready marketing graphics where chart labels must stay sharp. | Web thumbnails that need readable text without pixelation. | Automated report pipelines that embed charts as high‑resolution PNGs for PDF or slide decks.
// AI Prompts: Write C# code that exports an Aspose.Cells chart to a 600 DPI PNG while preserving localized axis labels. | Explain how to tune ImageOrPrintOptions for optimal file size versus resolution when rendering charts in Aspose.Cells. | Provide a script to batch‑export every chart in a workbook to PNG files with a custom DPI setting.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// Creates a workbook, adds sample data, builds a column chart, sets ImageOrPrintOptions to 300 DPI, and saves the chart as a high‑resolution PNG. The workbook can also be saved for reference.
class ExportChartHighResolution
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
        worksheet.Cells["B2"].PutValue(1200);
        worksheet.Cells["B3"].PutValue(800);
        worksheet.Cells["B4"].PutValue(1500);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Configure high‑resolution image options (e.g., 300 DPI)
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = ImageType.Png;         // Desired image format
        options.HorizontalResolution = 300;        // 300 DPI horizontal
        options.VerticalResolution = 300;          // 300 DPI vertical

        // Export the chart to a high‑resolution PNG file
        string outputImagePath = "HighResolutionChart.png";
        chart.ToImage(outputImagePath, options);

        // Save the workbook (optional, for reference)
        workbook.Save("ChartWorkbook.xlsx");
    }
}
