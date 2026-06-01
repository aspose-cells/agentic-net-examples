using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ChartPositionDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("Fruits");
                worksheet.Cells["A3"].PutValue("Vegetables");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(50);
                worksheet.Cells["B3"].PutValue(30);

                // Add a column chart at an initial position (rows 10‑20, columns 0‑5)
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 10, 0, 20, 5);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the chart data source
                chart.NSeries.Add("B2:B3", true);
                chart.NSeries.CategoryData = "A2:A3";

                // Move the chart to a new position while preserving its original size
                int topRow = 5;
                int leftColumn = 3;
                int bottomRow = topRow + (20 - 10);   // 15
                int rightColumn = leftColumn + (5 - 0); // 8
                chart.Move(topRow, leftColumn, bottomRow, rightColumn);

                // Save the workbook
                string outputPath = "ChartPositionDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}