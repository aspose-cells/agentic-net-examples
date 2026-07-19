// Title: Export a Waterfall Chart to a 300 DPI PNG with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills it with sample waterfall data, adds a Waterfall chart, sets ImageOrPrintOptions to PNG at 300 DPI, and uses chart.ToImage to generate a high‑resolution PNG file (WaterfallChart.png) for reports.
// Keywords: Aspose.Cells | C# | Waterfall chart | export PNG | 300 DPI | ImageOrPrintOptions | chart.ToImage | high resolution chart image | Aspose.Cells .NET | chart export | report graphics
// Common Searches: Aspose.Cells export waterfall chart PNG | C# save chart as 300 DPI PNG Aspose.Cells | How to set DPI when exporting charts in Aspose.Cells | Export chart to high resolution image .NET | Waterfall chart image export Aspose.Cells
// Developer Intent: Generate a 300 DPI PNG image of a Waterfall chart created with Aspose.Cells.
// Use Cases: Include a high‑resolution Waterfall chart in financial PDFs or PowerPoint decks. | Automate batch export of multiple worksheet charts to PNG for a BI dashboard. | Create printable chart graphics for corporate reports with precise DPI control.
// AI Prompts: Show me how to export any Aspose.Cells chart to a 600 DPI PNG using ImageOrPrintOptions. | Provide C# code to export a chart as JPEG with custom width, height, and DPI in Aspose.Cells. | Explain how to adjust chart size and resolution when calling ToImage in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// Creates a workbook, fills it with sample waterfall data, adds a Waterfall chart, sets ImageOrPrintOptions to PNG at 300 DPI, and uses chart.ToImage to generate a high‑resolution PNG file (WaterfallChart.png) for reports.
class ExportWaterfallChart
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the waterfall chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue("Start");
        worksheet.Cells["B2"].PutValue(100);
        worksheet.Cells["A3"].PutValue("Increase");
        worksheet.Cells["B3"].PutValue(30);
        worksheet.Cells["A4"].PutValue("Decrease");
        worksheet.Cells["B4"].PutValue(-20);
        worksheet.Cells["A5"].PutValue("End");
        worksheet.Cells["B5"].PutValue(110);

        // Add a Waterfall chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 15);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B5", true);
        chart.NSeries.CategoryData = "A2:A5";

        // Configure high‑resolution image options (300 DPI PNG)
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
        imgOptions.ImageType = ImageType.Png;
        imgOptions.HorizontalResolution = 300;
        imgOptions.VerticalResolution = 300;

        // Export the chart as a high‑resolution PNG image
        chart.ToImage("WaterfallChart.png", imgOptions);
    }
}
