// Title: Export a chart to a high‑resolution PNG (300 DPI) with Aspose.Cells for clear localized text
// Description: Demonstrates how to build a workbook, add sample data, create a column chart, configure ImageOrPrintOptions for 300 DPI PNG output, and save the chart as a high‑resolution image while optionally persisting the workbook.
// Keywords: Aspose.Cells C# export chart PNG 300 DPI | chart image resolution Aspose.Cells | high resolution chart rendering .NET | localized chart export Aspose.Cells | ImageOrPrintOptions DPI setting | Aspose.Cells chart to image
// Common Searches: Aspose.Cells export chart PNG 300 DPI C# | set DPI when exporting chart with Aspose.Cells | C# high resolution chart image Aspose.Cells | export localized chart as PNG using Aspose.Cells | ImageOrPrintOptions chart resolution .NET
// Developer Intent: Generate a high‑resolution PNG of a chart created with Aspose.Cells to verify that localized labels and data remain legible.
// Use Cases: Produce a 300 DPI PNG of a sales chart for inclusion in marketing collateral. | Validate axis and data‑label clarity after applying language‑specific fonts. | Automate UI tests that compare rendered chart images against baseline screenshots.
// AI Prompts: Show me C# code to export an Aspose.Cells chart as a 300 DPI PNG while preserving localized fonts. | How can I change ImageOrPrintOptions to 600 DPI and save the chart as a TIFF in Aspose.Cells? | Explain how to adjust chart dimensions and DPI when rendering a high‑resolution image with Aspose.Cells .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// Demonstrates how to build a workbook, add sample data, create a column chart, configure ImageOrPrintOptions for 300 DPI PNG output, and save the chart as a high‑resolution image while optionally persisting the workbook.
class ExportHighResChart
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

        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(1200);
        sheet.Cells["B3"].PutValue(800);
        sheet.Cells["B4"].PutValue(1500);

        // Add a column chart
        int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIdx];

        // Set the data range for the chart
        chart.SetChartDataRange("A1:B4", true);

        // Configure high‑resolution image options (e.g., 300 DPI)
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
        {
            ImageType = ImageType.Png,
            HorizontalResolution = 300,
            VerticalResolution = 300
        };

        // Export the chart to a high‑resolution PNG file
        string outputPath = "HighResChart.png";
        chart.ToImage(outputPath, imgOptions);

        // Optionally save the workbook for reference
        workbook.Save("ChartWorkbook.xlsx");

        Console.WriteLine($"Chart exported to high‑resolution PNG at: {outputPath}");
    }
}
