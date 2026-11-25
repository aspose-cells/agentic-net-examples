// Title: Export a Japanese Excel chart to a JPEG file using Aspose.Cells C# ToImage method
// Common Searches: C# Aspose.Cells export Excel chart with Japanese labels to JPEG | How to save a localized chart as an image using Aspose.Cells ToImage | Aspose.Cells chart ToImage method example for JPEG output | Generate JPEG from Excel column chart containing Unicode text in C# | Export Excel chart with non‑English title to JPEG using Aspose.Cells
// Tags: chart to jpeg conversion Aspose.Cells | export localized chart image c# | ToImage method Aspose.Cells jpeg | Japanese chart rendering Aspose.Cells | column chart image export C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// The sample creates a workbook, fills cells with Japanese category and sales data, adds a column chart with a Japanese title, and uses the chart's ToImage method to export the chart as a JPEG file named "LocalizedChart.jpg" in the current directory.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate worksheet with localized (Japanese) data
            sheet.Cells["A1"].PutValue("カテゴリ"); // Category
            sheet.Cells["A2"].PutValue("りんご");   // Apple
            sheet.Cells["A3"].PutValue("オレンジ"); // Orange
            sheet.Cells["A4"].PutValue("バナナ");   // Banana

            sheet.Cells["B1"].PutValue("売上"); // Sales
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["B3"].PutValue(800);
            sheet.Cells["B4"].PutValue(1500);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Define the data range for the chart
            chart.SetChartDataRange("A1:B4", true);

            // Set a localized title for the chart
            chart.Title.Text = "売上チャート"; // "Sales Chart"

            // Export the chart to a JPEG image using the ToImage method
            string outputFile = Path.Combine(Environment.CurrentDirectory, "LocalizedChart.jpg");
            chart.ToImage(outputFile, ImageType.Jpeg);

            Console.WriteLine($"Chart successfully exported to: {outputFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
