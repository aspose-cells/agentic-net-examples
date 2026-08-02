using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class ExportChartAsPng
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

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.SetChartDataRange("A1:B4", true);

        // Ensure the output directory exists
        string outputFolder = "output";
        Directory.CreateDirectory(outputFolder);

        // Define the full path for the PNG image
        string imagePath = Path.Combine(outputFolder, "ChartImage.png");

        // Export the chart as a PNG image
        chart.ToImage(imagePath, ImageType.Png);

        // (Optional) Save the workbook for reference
        workbook.Save(Path.Combine(outputFolder, "Workbook.xlsx"));
    }
}