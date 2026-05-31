using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class ChartThumbnailGenerator
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

        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["B2"].PutValue(1200);
        sheet.Cells["B3"].PutValue(800);
        sheet.Cells["B4"].PutValue(1500);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.SetChartDataRange("A1:B4", true);

        // Configure image options to produce a 200x150 thumbnail
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
        {
            ImageType = ImageType.Png,
            // Keep aspect ratio false to force exact dimensions
        };
        imgOptions.SetDesiredSize(200, 150, false);

        // Save the chart as an image file (thumbnail)
        chart.ToImage("chart_thumbnail.png", imgOptions);

        // Optional: Save the workbook if needed
        workbook.Save("chart_workbook.xlsx");
    }
}