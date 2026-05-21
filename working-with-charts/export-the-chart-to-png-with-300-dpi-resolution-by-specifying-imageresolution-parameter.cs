using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class ExportChartWithDpi
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
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["B2"].PutValue(1200);
        sheet.Cells["B3"].PutValue(800);
        sheet.Cells["B4"].PutValue(1500);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Configure image options: PNG format with 300 DPI resolution
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Png,
            HorizontalResolution = 300,
            VerticalResolution = 300
        };

        // Export the chart to a PNG file using the specified resolution
        chart.ToImage("Chart_300dpi.png", options);

        Console.WriteLine("Chart exported to 'Chart_300dpi.png' with 300 DPI resolution.");
    }
}