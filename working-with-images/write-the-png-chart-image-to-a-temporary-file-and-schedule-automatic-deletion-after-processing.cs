using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class Program
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
        worksheet.Cells["B2"].PutValue(1200);
        worksheet.Cells["B3"].PutValue(800);
        worksheet.Cells["B4"].PutValue(1500);

        // Add a column chart
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.SetChartDataRange("A1:B4", true);

        // Create a temporary file name with .png extension
        string tempFilePath = Path.ChangeExtension(Path.GetTempFileName(), ".png");

        try
        {
            // Save the chart as a PNG image to the temporary file
            chart.ToImage(tempFilePath, ImageType.Png);

            // At this point you can process the image as needed.
            // For demonstration, we just output the file path and size.
            FileInfo info = new FileInfo(tempFilePath);
            Console.WriteLine($"Chart image saved to temporary file: {tempFilePath}");
            Console.WriteLine($"File size: {info.Length} bytes");
        }
        finally
        {
            // Ensure the temporary file is deleted after processing
            if (File.Exists(tempFilePath))
            {
                try
                {
                    File.Delete(tempFilePath);
                    Console.WriteLine("Temporary chart image deleted.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to delete temporary file: {ex.Message}");
                }
            }
        }
    }
}