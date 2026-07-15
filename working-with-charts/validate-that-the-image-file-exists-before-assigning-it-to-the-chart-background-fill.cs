using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class SetChartBackgroundImage
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];
        chart.SetChartDataRange("A1:B4", true);

        // Path to the image that will be used as chart background
        string imagePath = "chartBackground.png";

        // Validate that the image file exists before assigning it
        if (File.Exists(imagePath))
        {
            // Read the image file into a byte array
            byte[] imageData = File.ReadAllBytes(imagePath);

            // Set the fill type to texture and assign the image data
            chart.ChartArea.Area.FillFormat.FillType = FillType.Texture;
            chart.ChartArea.Area.FillFormat.TextureFill.ImageData = imageData;
        }
        else
        {
            Console.WriteLine($"Image file not found: {imagePath}");
        }

        // Save the workbook with the chart
        workbook.Save("ChartWithBackground.xlsx");
    }
}