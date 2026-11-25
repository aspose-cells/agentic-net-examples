// Title: How to retrieve and log chart position and size (including PlotArea) using Aspose.Cells for .NET C#
// Common Searches: Aspose.Cells get chart X and Y coordinates in C# | C# code to read chart object position points with Aspose.Cells | Retrieve plot area dimensions of a chart using Aspose.Cells .NET | Log chart placement coordinates and size in Aspose.Cells for .NET | How to obtain chart shape size in points with Aspose.Cells C#
// Tags: Aspose.Cells chart object position | Aspose.Cells retrieve chart coordinates | Aspose.Cells plot area dimensions | Aspose.Cells chart shape size points | Aspose.Cells C# chart placement debugging

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace ChartPositionLogger
{
    // The example creates a workbook, adds sample data and a column chart, then accesses the ChartObject to obtain its X, Y, width, and height values (in points). It also reads the PlotArea's X, Y, width, and height, logs all coordinates to the console for debugging, and saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Retrieve chart position coordinates via ChartObject
                ChartShape chartShape = chart.ChartObject;

                double chartX = chartShape.X;          // X position (in points)
                double chartY = chartShape.Y;          // Y position (in points)
                double chartWidth = chartShape.Width;  // Width (in points)
                double chartHeight = chartShape.Height;// Height (in points)

                // Log the coordinates for debugging
                Console.WriteLine("Chart Position and Size (ChartObject):");
                Console.WriteLine($"  X      = {chartX} points");
                Console.WriteLine($"  Y      = {chartY} points");
                Console.WriteLine($"  Width  = {chartWidth} points");
                Console.WriteLine($"  Height = {chartHeight} points");

                // Additionally, log PlotArea coordinates if needed
                PlotArea plotArea = chart.PlotArea;
                Console.WriteLine("\nPlotArea Position and Size:");
                Console.WriteLine($"  X      = {plotArea.X} points");
                Console.WriteLine($"  Y      = {plotArea.Y} points");
                Console.WriteLine($"  Width  = {plotArea.Width} points");
                Console.WriteLine($"  Height = {plotArea.Height} points");

                // Save the workbook (optional, just to verify that everything works)
                string outputPath = "ChartPositionDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"\nWorkbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine("An error occurred while retrieving chart coordinates:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
