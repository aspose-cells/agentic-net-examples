using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class WaterfallChartAxisNumberFormat
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add sample data for a waterfall chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("Start");
                worksheet.Cells["A3"].PutValue("Increase");
                worksheet.Cells["A4"].PutValue("Decrease");
                worksheet.Cells["A5"].PutValue("End");

                worksheet.Cells["B1"].PutValue("Amount");
                worksheet.Cells["B2"].PutValue(50000);
                worksheet.Cells["B3"].PutValue(20000);
                worksheet.Cells["B4"].PutValue(-15000);
                worksheet.Cells["B5"].PutValue(55000);

                // Add a Waterfall chart
                int chartIndex = worksheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 15);
                Chart chart = worksheet.Charts[chartIndex];

                // Set data range for the chart
                chart.NSeries.Add("B2:B5", true);          // Values
                chart.NSeries.CategoryData = "A2:A5";      // Categories

                // Set custom number format for the value axis tick labels to display currency
                chart.ValueAxis.TickLabels.NumberFormat = "$#,##0";

                // Save the workbook
                string outputPath = "WaterfallChartAxisCurrencyFormat.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            WaterfallChartAxisNumberFormat.Run();
        }
    }
}