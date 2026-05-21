using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

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

        // Configure high‑resolution JPEG export options
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = ImageType.Jpeg;        // JPEG format
        options.HorizontalResolution = 300;        // 300 DPI horizontal
        options.VerticalResolution = 300;          // 300 DPI vertical
        options.Quality = 90;                      // JPEG quality (0‑100)

        // Export the chart as a high‑resolution JPEG image
        chart.ToImage("HighResChart.jpg", options);

        // (Optional) Save the workbook for reference
        workbook.Save("ChartWorkbook.xlsx");
    }
}