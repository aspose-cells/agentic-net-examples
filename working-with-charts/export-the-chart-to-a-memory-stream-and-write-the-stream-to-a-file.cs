using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

public class ExportChartToFile
{
    public static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.SetChartDataRange("A1:B4", true);

        // Export the chart to a memory stream and then write the stream to a file
        using (MemoryStream imageStream = new MemoryStream())
        {
            // Convert the chart to PNG image and write it into the stream
            chart.ToImage(imageStream, ImageType.Png);

            // Ensure the stream position is at the beginning before reading
            imageStream.Position = 0;

            // Write the stream contents to a file
            File.WriteAllBytes("ChartImage.png", imageStream.ToArray());

            Console.WriteLine("Chart image saved to ChartImage.png");
        }

        // Optionally save the workbook (demonstrating the Save rule)
        workbook.Save("ChartWorkbook.xlsx");
    }
}