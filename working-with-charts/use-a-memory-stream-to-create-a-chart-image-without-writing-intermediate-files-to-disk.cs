using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class ChartToImageMemoryStream
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data source for the chart
        chart.NSeries.Add("B2:B3", true);
        chart.NSeries.CategoryData = "A2:A3";

        // Create a MemoryStream using the provided factory rule
        CustomImplementationFactory factory = new CustomImplementationFactory();
        using (MemoryStream stream = factory.CreateMemoryStream())
        {
            // Render the chart to the stream in PNG format using the Chart.ToImage rule
            chart.ToImage(stream, ImageType.Png);

            // Reset the stream position if further reading is required
            stream.Position = 0;

            // Example: obtain the image bytes (no file is written to disk)
            byte[] imageBytes = stream.ToArray();
            Console.WriteLine($"Chart image generated in memory. Byte size: {imageBytes.Length}");
        }
    }
}