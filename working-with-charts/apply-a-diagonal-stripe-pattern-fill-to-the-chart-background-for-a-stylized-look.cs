using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartPatternFill
{
    public class ApplyDiagonalStripePattern
    {
        public static void Run()
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

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Apply a diagonal stripe pattern fill to the chart area background
                chart.ChartArea.Area.FillFormat.FillType = FillType.Pattern; // Use pattern fill
                chart.ChartArea.Area.FillFormat.PatternFill.Pattern = FillPattern.LightDownwardDiagonal; // Diagonal stripe
                chart.ChartArea.Area.FillFormat.PatternFill.ForegroundColor = Color.LightBlue; // Foreground color
                chart.ChartArea.Area.FillFormat.PatternFill.BackgroundColor = Color.DarkBlue; // Background color

                // Save the workbook
                string outputPath = "ChartWithDiagonalStripePattern.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ApplyDiagonalStripePattern.Run();
        }
    }
}