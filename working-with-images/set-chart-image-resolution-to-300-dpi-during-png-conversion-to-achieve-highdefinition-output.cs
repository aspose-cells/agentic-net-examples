using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class SetChartResolutionExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["A4"].PutValue("Banana");

        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(80);
        sheet.Cells["B4"].PutValue(150);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Configure image options for high‑definition PNG output
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Png,          // Ensure PNG format
            HorizontalResolution = 300,         // 300 DPI horizontally
            VerticalResolution = 300            // 300 DPI vertically
        };

        // Save the chart as a PNG image with the specified DPI
        string outputPath = "HighResolutionChart.png";
        chart.ToImage(outputPath, options);

        Console.WriteLine($"Chart saved to '{outputPath}' with 300 DPI resolution.");
    }
}