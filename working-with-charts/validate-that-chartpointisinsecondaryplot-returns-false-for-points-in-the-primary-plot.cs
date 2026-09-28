// Title: Check that ChartPoint.IsInSecondaryPlot returns false for all points in a primary‑only Pie chart using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that creates a pie chart, loops through each series point, and prints whether ChartPoint.IsInSecondaryPlot is false. | Generate a .NET example that asserts the IsInSecondaryPlot property is false for every chart point in a chart that has only a primary plot.
// Common Searches: Aspose.Cells how to determine if a chart point belongs to secondary plot in C# | C# verify ChartPoint.IsInSecondaryPlot returns false for pie chart series | example code to iterate chart points and check IsInSecondaryPlot property using Aspose.Cells .NET | validate primary plot points in Aspose.Cells chart with IsInSecondaryPlot
// Tags: Aspose.Cells verify primary plot points | C# chart point IsInSecondaryPlot false | Aspose.Cells pie chart series iteration | Aspose.Cells .NET chart validation

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ChartPointSecondaryPlotValidation
{
    // The sample creates a workbook, adds sample data, builds a pie chart, iterates through the chart series points, confirms that ChartPoint.IsInSecondaryPlot is false for each point, outputs the results, and saves the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["A5"].PutValue("D");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(40);
                sheet.Cells["B3"].PutValue(30);
                sheet.Cells["B4"].PutValue(20);
                sheet.Cells["B5"].PutValue(10);

                // Add a Pie chart (primary plot only; secondary plot not applicable)
                int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart series
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Validate that no points are in the secondary plot
                bool allPrimary = true;
                for (int i = 0; i < chart.NSeries[0].Points.Count; i++)
                {
                    ChartPoint point = chart.NSeries[0].Points[i];
                    if (point.IsInSecondaryPlot)
                    {
                        allPrimary = false;
                        Console.WriteLine($"Point {i} is incorrectly marked as secondary.");
                    }
                    else
                    {
                        Console.WriteLine($"Point {i} correctly reports IsInSecondaryPlot = false.");
                    }
                }

                Console.WriteLine(allPrimary
                    ? "All points are correctly identified as primary plot."
                    : "Some points are incorrectly identified as secondary plot.");

                // Save the workbook (optional, just to complete the lifecycle)
                string outputPath = "ChartPointSecondaryPlotValidation.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
