using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class Program
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
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.SetChartDataRange("A1:B4", true);

        // Configure image options: PNG format with 300 DPI resolution
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
        imgOptions.ImageType = ImageType.Png;
        imgOptions.HorizontalResolution = 300;
        imgOptions.VerticalResolution = 300;

        // Define output paths (same folder)
        string outputFolder = Directory.GetCurrentDirectory();
        string workbookPath = Path.Combine(outputFolder, "ChartWorkbook.xlsx");
        string chartImagePath = Path.Combine(outputFolder, "ChartImage.png");

        // Save the workbook
        workbook.Save(workbookPath);

        // Export the chart as a PNG image with the specified DPI
        chart.ToImage(chartImagePath, imgOptions);

        Console.WriteLine($"Workbook saved to: {workbookPath}");
        Console.WriteLine($"Chart image saved to: {chartImagePath}");
    }
}