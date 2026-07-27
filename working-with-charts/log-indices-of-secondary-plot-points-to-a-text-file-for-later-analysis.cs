using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class LogSecondaryPlotPointsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for a pie chart
                worksheet.Cells["A1"].PutValue(50);
                worksheet.Cells["A2"].PutValue(100);
                worksheet.Cells["A3"].PutValue(150);
                worksheet.Cells["B1"].PutValue(60);
                worksheet.Cells["B2"].PutValue(32);
                worksheet.Cells["B3"].PutValue(50);

                // Add a pie chart that supports secondary plots
                int chartIndex = worksheet.Charts.Add(ChartType.Pie, 5, 0, 25, 10);
                Chart chart = worksheet.Charts[chartIndex];
                chart.NSeries.Add("A1:B3", true);

                // Mark some points as belonging to the secondary plot (if they exist)
                Series series = chart.NSeries[0];
                if (series.Points.Count > 2)
                    series.Points[2].IsInSecondaryPlot = true; // third point (index 2)
                if (series.Points.Count > 3)
                    series.Points[3].IsInSecondaryPlot = true; // fourth point (index 3)

                // Collect indices of points that are in the secondary plot
                List<int> secondaryIndices = new List<int>();
                for (int i = 0; i < series.Points.Count; i++)
                {
                    if (series.Points[i].IsInSecondaryPlot)
                        secondaryIndices.Add(i);
                }

                // Write the collected indices to a text file
                string logFilePath = "SecondaryPlotPoints.txt";
                using (StreamWriter writer = new StreamWriter(logFilePath, false))
                {
                    writer.WriteLine("Indices of secondary plot points:");
                    foreach (int idx in secondaryIndices)
                        writer.WriteLine(idx);
                }

                // Save the workbook with the chart
                workbook.Save("PieChartWithSecondaryPlot.xlsx");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            LogSecondaryPlotPointsDemo.Run();
        }
    }
}