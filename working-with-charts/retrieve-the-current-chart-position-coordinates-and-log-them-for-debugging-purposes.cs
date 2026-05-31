using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartDebug
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
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

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";      // Categories

                // Calculate the chart so that shape positions are populated
                chart.Calculate();

                // Retrieve the first point of the first series
                ChartPoint point = chart.NSeries[0].Points[0];

                // Log the chart point's position in pixels
                Console.WriteLine("Chart Point Position (pixels):");
                Console.WriteLine($"  ShapeXPx = {point.ShapeXPx}");
                Console.WriteLine($"  ShapeYPx = {point.ShapeYPx}");

                // Log the chart object's top‑left and lower‑right cell coordinates
                Console.WriteLine("Chart Object Position (cell indices):");
                Console.WriteLine($"  UpperLeftRow    = {chart.ChartObject.UpperLeftRow}");
                Console.WriteLine($"  UpperLeftColumn = {chart.ChartObject.UpperLeftColumn}");
                Console.WriteLine($"  LowerRightRow   = {chart.ChartObject.LowerRightRow}");
                Console.WriteLine($"  LowerRightColumn= {chart.ChartObject.LowerRightColumn}");

                // Save the workbook (debugging purpose)
                string outputPath = "ChartDebugOutput.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}