// Title: Export a Waterfall chart to a high‑resolution PNG file using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export waterfall chart as PNG with custom resolution | How to save an Excel waterfall chart to a PNG image in .NET | C# code for rendering Aspose.Cells chart to high‑resolution PNG | Generate PNG image from Waterfall chart using Aspose.Cells library | Set image size when exporting chart to PNG with Aspose.Cells C#
// Tags: waterfall chart export to PNG Aspose.Cells | Aspose.Cells ToImage high resolution chart | C# render Excel chart as image | save Aspose.Cells chart as PNG file | chart image generation .NET Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, fills it with waterfall chart data, adds a Waterfall chart, sets its title, and then uses the Chart.ToImage method to render the chart directly to a high‑resolution PNG file named "WaterfallChart.png", handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data for the waterfall chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Start");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["A3"].PutValue("Revenue");
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["A4"].PutValue("Cost");
            sheet.Cells["B4"].PutValue(-70);
            sheet.Cells["A5"].PutValue("Profit");
            sheet.Cells["B5"].PutValue(80);
            sheet.Cells["A6"].PutValue("End");
            sheet.Cells["B6"].PutValue(160);

            // Add a waterfall chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 10, 0, 30, 10);
            Chart waterfallChart = sheet.Charts[chartIndex];

            // Set the series data range and categories
            waterfallChart.NSeries.Add("B2:B6", true);
            waterfallChart.NSeries.CategoryData = "A2:A6";

            // Optional: set chart title
            waterfallChart.Title.Text = "Waterfall Chart Example";

            // Render the chart to an image file
            string outputPath = "WaterfallChart.png";
            try
            {
                // Save chart directly to file (default format is PNG)
                waterfallChart.ToImage(outputPath);
                Console.WriteLine($"Waterfall chart image saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception renderEx)
            {
                Console.WriteLine($"Error during chart rendering: {renderEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
