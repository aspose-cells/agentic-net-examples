using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsWaterfallDemo
{
    public class WaterfallChartWithTotal
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the waterfall chart
                // Column A – Categories, Column B – Values
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Start");
                sheet.Cells["A3"].PutValue("Increase 1");
                sheet.Cells["A4"].PutValue("Increase 2");
                sheet.Cells["A5"].PutValue("Decrease");
                sheet.Cells["A6"].PutValue("Total");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(100);   // Starting value
                sheet.Cells["B3"].PutValue(30);    // Positive change
                sheet.Cells["B4"].PutValue(20);    // Positive change
                sheet.Cells["B5"].PutValue(-15);   // Negative change
                sheet.Cells["B6"].PutValue(0);     // Placeholder for total (will be marked as total)

                // Add a Waterfall chart
                int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data series (values) and category (labels)
                chart.NSeries.Add("B2:B6", true);
                chart.NSeries.CategoryData = "A2:A6";

                // Mark the last data point (index 4, zero‑based) as a total
                Series series = chart.NSeries[0];
                series.LayoutProperties.Subtotals = new int[] { 4 };

                // Optional: give the chart a title
                chart.Title.Text = "Waterfall Chart with Total";

                // Define output file path
                string outputPath = "WaterfallChartWithTotal.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            WaterfallChartWithTotal.Run();
        }
    }
}