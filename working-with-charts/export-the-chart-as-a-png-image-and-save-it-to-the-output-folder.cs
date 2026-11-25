// Title: Export an Aspose.Cells column chart to a PNG file and store it in an output folder with C#
// Common Searches: C# Aspose.Cells export column chart to PNG file | How to save Aspose.Cells chart as an image in .NET | Aspose.Cells chart ToImage method example for PNG output | Create output directory and write chart image using Aspose.Cells | Export Excel chart to PNG using Aspose.Cells API in C# | Save generated chart image to file system with Aspose.Cells
// Tags: Aspose.Cells chart ToImage PNG | C# column chart image generation | export Excel chart to PNG using Aspose | output folder creation for chart export .NET | Aspose.Cells chart rendering to file

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace ExportChartToPng
{
    // The sample creates a workbook with fruit‑sales data, adds a column chart, and uses the Chart.ToImage method to render the chart as a PNG file. It ensures an "output" directory exists, writes the PNG image there, and also saves the workbook as an XLSX file for reference.
    class Program
    {
        static void Main(string[] args)
        {
            // Define paths
            string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "output");
            string chartImagePath = Path.Combine(outputFolder, "Chart.png");

            try
            {
                // Ensure the output directory exists
                if (!Directory.Exists(outputFolder))
                {
                    Directory.CreateDirectory(outputFolder);
                    Console.WriteLine($"Created output directory: {outputFolder}");
                }

                // Create a new workbook and populate it with sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["A3"].PutValue("Orange");
                sheet.Cells["A4"].PutValue("Banana");

                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["B2"].PutValue(1200);
                sheet.Cells["B3"].PutValue(800);
                sheet.Cells["B4"].PutValue(1500);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Optional: give the chart a title
                chart.Title.Text = "Fruit Sales";

                // Export the chart to PNG using the direct ToImage overload
                chart.ToImage(chartImagePath, ImageType.Png);
                Console.WriteLine($"Chart exported successfully to: {chartImagePath}");

                // If you also want to save the workbook for reference
                string workbookPath = Path.Combine(outputFolder, "SampleWorkbook.xlsx");
                workbook.Save(workbookPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to: {workbookPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error occurred: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
