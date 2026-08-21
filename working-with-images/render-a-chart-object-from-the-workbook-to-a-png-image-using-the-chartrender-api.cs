// Title: C# Example: Export an Aspose.Cells Chart to PNG using Chart.ToImage API
// Description: Creates a workbook, fills A1:B4 with categories and values, adds a column chart, binds it to the range, and uses Chart.ToImage with ImageType.Png to save the chart as "ChartImage.png" while optionally saving the workbook.
// Keywords: Aspose.Cells | Chart.ToImage | C# chart export PNG | export Excel chart as image | Aspose.Cells Chart API | save chart to PNG | render chart to image | Aspose.Cells .NET | ImageType.Png
// Common Searches: Aspose.Cells export chart to PNG C# | Chart.ToImage example .NET | How to save Excel chart as image using Aspose | C# render Excel chart to PNG | Aspose.Cells chart image generation
// Developer Intent: Generate a PNG file from a chart built in an Aspose.Cells workbook.
// Use Cases: Automate creation of chart images for web dashboards without Excel | Produce printable chart graphics for reports directly from server‑side code | Batch‑process multiple charts in a workbook and store each as a PNG file
// AI Prompts: Write C# code that creates a pie chart from range C1:D5 and saves it as a JPEG using Aspose.Cells. | Explain how to set chart size, font, and colors before calling Chart.ToImage. | Provide a loop that iterates through all worksheets and renders every chart to separate PNG files.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace ChartRenderDemo
{
    // Creates a workbook, fills A1:B4 with categories and values, adds a column chart, binds it to the range, and uses Chart.ToImage with ImageType.Png to save the chart as "ChartImage.png" while optionally saving the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
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

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.SetChartDataRange("A1:B4", true);

            // Render the chart to a PNG image using the Chart ToImage API
            string imagePath = "ChartImage.png";
            chart.ToImage(imagePath, ImageType.Png);

            Console.WriteLine($"Chart has been rendered to PNG image at: {imagePath}");

            // Optionally, save the workbook for reference
            workbook.Save("ChartWorkbook.xlsx");
        }
    }
}
