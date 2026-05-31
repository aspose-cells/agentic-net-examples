using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class ExportChartHighDPI
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
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data source for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Configure image options for high‑resolution PNG (300 DPI)
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = ImageType.Png;
        options.HorizontalResolution = 300; // 300 DPI horizontally
        options.VerticalResolution = 300;   // 300 DPI vertically

        // Export the chart to a PNG file using the high‑DPI settings
        chart.ToImage("ChartHighDPI.png", options);
    }
}