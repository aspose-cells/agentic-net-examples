using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["A3"].PutValue("Banana");
        worksheet.Cells["A4"].PutValue("Cherry");

        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(120);
        worksheet.Cells["B3"].PutValue(80);
        worksheet.Cells["B4"].PutValue(150);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Configure image options for high‑definition PNG (300 DPI)
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = ImageType.Png;
        options.HorizontalResolution = 300; // 300 DPI horizontally
        options.VerticalResolution = 300;   // 300 DPI vertically

        // Export the chart to a PNG file using the specified DPI settings
        chart.ToImage("Chart_300dpi.png", options);
    }
}