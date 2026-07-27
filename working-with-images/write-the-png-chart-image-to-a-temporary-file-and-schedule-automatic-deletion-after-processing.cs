using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class ChartToTempPng
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["A4"].PutValue("Banana");
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["B2"].PutValue(1200);
        sheet.Cells["B3"].PutValue(800);
        sheet.Cells["B4"].PutValue(1500);

        // Add a column chart to the worksheet
        int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIdx];

        // Set the data range for the chart
        chart.SetChartDataRange("A1:B4", true);

        // Create a temporary file name with .png extension
        string tempPngPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".png");

        try
        {
            // Save the chart as a PNG image to the temporary file
            // Uses the Chart.ToImage(string, ImageType) overload
            chart.ToImage(tempPngPath, ImageType.Png);

            // At this point you can process the PNG file as needed.
            // For demonstration, we just output its size.
            FileInfo info = new FileInfo(tempPngPath);
            Console.WriteLine($"Chart image saved to temporary file: {tempPngPath}");
            Console.WriteLine($"File size: {info.Length} bytes");
        }
        finally
        {
            // Ensure the temporary file is deleted after processing
            if (File.Exists(tempPngPath))
            {
                try
                {
                    File.Delete(tempPngPath);
                    Console.WriteLine("Temporary file deleted.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to delete temporary file: {ex.Message}");
                }
            }
        }
    }
}