// Title: Export an Aspose.Cells pie chart to a JPEG image using Chart.ToImage in C#
// Common Searches: C# Aspose.Cells export pie chart as JPEG file | How to use Chart.ToImage to save Excel chart in JPEG format with Aspose.Cells | Aspose.Cells Chart.ToImage JPEG example for .NET | Convert Excel pie chart to JPEG image using Aspose.Cells in C# | Saving Aspose.Cells chart to JPEG without System.Drawing
// Tags: Aspose.Cells export chart to JPEG | Chart.ToImage JPEG format C# | pie chart image generation Aspose.Cells | Excel chart to JPEG conversion .NET | Aspose.Cells image export without System.Drawing | C# Aspose.Cells chart image options

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ExportPieChartExample
{
    // The sample creates a new workbook, adds a pie chart with fruit data, sets its series and title, and demonstrates exporting the chart directly to a JPEG image using the Chart.ToImage method with the JPEG format. The workbook is then saved as an XLSX file.
    class ExportPieChartToJpeg
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate data for the pie chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(50);
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(30);
                sheet.Cells["A4"].PutValue("Cherry");
                sheet.Cells["B4"].PutValue(20);

                // Add a pie chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 15, 5);
                Chart pieChart = sheet.Charts[chartIndex];

                // Set the series values and categories
                pieChart.NSeries.Add("B2:B4", true);
                pieChart.NSeries.CategoryData = "A2:A4";

                // Optional: set a title for the chart
                pieChart.Title.Text = "Fruit Distribution";

                // Export the chart to an image file (PNG format, no System.Drawing dependency)
                pieChart.ToImage("pieChart.png");

                // Save the workbook
                workbook.Save("WorkbookWithPieChart.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
