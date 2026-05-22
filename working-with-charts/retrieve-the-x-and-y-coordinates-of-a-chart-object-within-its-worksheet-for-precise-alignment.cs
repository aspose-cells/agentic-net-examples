using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartPositionDemo
{
    public class RetrieveChartCoordinates
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
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

            // Add a column chart (rows 5‑20, columns 0‑8)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Retrieve the chart's shape object
            ChartShape chartShape = chart.ChartObject;

            // X and Y coordinates of the chart's upper‑left corner in pixels
            int chartXPixel = chartShape.X;
            int chartYPixel = chartShape.Y;

            // Upper‑left cell indices where the chart starts
            int upperLeftRow = chartShape.UpperLeftRow;
            int upperLeftColumn = chartShape.UpperLeftColumn;

            // Output the retrieved coordinates
            Console.WriteLine($"Chart upper‑left corner (pixels): X = {chartXPixel}, Y = {chartYPixel}");
            Console.WriteLine($"Chart starts at cell: Row = {upperLeftRow}, Column = {upperLeftColumn}");

            // Save the workbook
            string outputPath = "ChartPositionDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
    }
}