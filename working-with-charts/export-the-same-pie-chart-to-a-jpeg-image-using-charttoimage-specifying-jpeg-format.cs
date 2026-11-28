// Title: Export a pie chart to a JPEG file using Aspose.Cells Chart.ToImage in C#
// Common Searches: C# Aspose.Cells export pie chart as JPEG image | How to save an Excel pie chart to JPEG with Chart.ToImage | Aspose.Cells Chart.ToImage method example for JPEG output | Generate JPEG file from a pie chart in Aspose.Cells C# | Export Excel chart to JPEG using Aspose.Cells library
// Tags: Aspose.Cells Chart.ToImage JPEG export | C# export Excel pie chart to JPEG | Aspose.Cells generate chart image | Chart image conversion JPEG C# | Aspose.Cells chart export as JPEG

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace ExportPieChartToJpeg
{
    // The sample creates a workbook, fills cells A1:B4 with category and value data, adds a pie chart, sets a title, and then uses Chart.ToImage with ImageType.Jpeg to export the chart as a JPEG file named "PieChartExport.jpg", handling any exceptions that may occur.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pie chart
                // Category column (A) and values column (B)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["A3"].PutValue("Orange");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("Banana");
                sheet.Cells["B4"].PutValue(50);

                // Add a pie chart to the worksheet
                // Parameters: chart type, upper left row, upper left column, lower right row, lower right column
                int chartIndex = sheet.Charts.Add(ChartType.Pie, 6, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart (including category labels)
                // The range includes both categories and values
                chart.SetChartDataRange("A1:B4", true);

                // Optionally, set a title for the chart
                chart.Title.Text = "Fruit Sales Distribution";

                // Export the chart to a JPEG image using the overload that specifies ImageType
                string outputImagePath = "PieChartExport.jpg";
                chart.ToImage(outputImagePath, ImageType.Jpeg);

                Console.WriteLine($"Pie chart successfully exported to JPEG image: {outputImagePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during chart creation or export: {ex.Message}");
            }
        }
    }
}
