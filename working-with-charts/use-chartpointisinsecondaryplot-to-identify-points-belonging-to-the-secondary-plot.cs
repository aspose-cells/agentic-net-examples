using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class ChartPointSecondaryPlotDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a pie chart
                sheet.Cells["A1"].PutValue(50);
                sheet.Cells["A2"].PutValue(100);
                sheet.Cells["A3"].PutValue(150);
                sheet.Cells["B1"].PutValue(60);
                sheet.Cells["B2"].PutValue(32);
                sheet.Cells["B3"].PutValue(50);

                // Add a pie chart that supports secondary plots
                int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("A1:B3", true);

                // Mark specific points as belonging to the secondary plot
                // Only the third point exists (index 2) in this series
                if (chart.NSeries[0].Points.Count > 2)
                {
                    chart.NSeries[0].Points[2].IsInSecondaryPlot = true;
                }

                // Iterate through all points and identify those in the secondary plot
                Console.WriteLine("Points belonging to the secondary plot:");
                for (int i = 0; i < chart.NSeries[0].Points.Count; i++)
                {
                    ChartPoint point = chart.NSeries[0].Points[i];
                    if (point.IsInSecondaryPlot)
                    {
                        // Output the index and its Y value for demonstration
                        Console.WriteLine($"  Point Index: {i}, YValue: {point.YValue}");
                    }
                }

                // Save the workbook to an Excel file
                string outputPath = "ChartPointSecondaryPlotDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
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
            ChartPointSecondaryPlotDemo.Run();
        }
    }
}