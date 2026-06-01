using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

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
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(15);
        sheet.Cells["B4"].PutValue(7);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Configure image options: PNG format with 300 DPI resolution
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
        imgOptions.ImageType = ImageType.Png;
        imgOptions.HorizontalResolution = 300;
        imgOptions.VerticalResolution = 300;

        // Export the chart as a PNG image using the configured options
        string chartImagePath = "ChartImage.png";
        chart.ToImage(chartImagePath, imgOptions);

        // Save the workbook in the same directory
        string workbookPath = "WorkbookWithChart.xlsx";
        workbook.Save(workbookPath);

        Console.WriteLine($"Chart exported to '{chartImagePath}' with 300 DPI.");
        Console.WriteLine($"Workbook saved to '{workbookPath}'.");
    }
}