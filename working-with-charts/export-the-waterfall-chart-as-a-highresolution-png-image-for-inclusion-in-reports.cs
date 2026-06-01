using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

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
        int chartIndex = worksheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 10);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.SetChartDataRange("A1:B5", true);
        chart.Title.Text = "Waterfall Chart Example";

        // Configure high‑resolution image options (300 DPI PNG)
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
        imgOptions.ImageType = ImageType.Png;
        imgOptions.HorizontalResolution = 300;
        imgOptions.VerticalResolution = 300;

        // Export the chart to a high‑resolution PNG file
        string outputFile = "WaterfallChart.png";
        chart.ToImage(outputFile, imgOptions);

        Console.WriteLine($"Waterfall chart exported successfully to '{outputFile}' with 300 DPI.");
    }
}