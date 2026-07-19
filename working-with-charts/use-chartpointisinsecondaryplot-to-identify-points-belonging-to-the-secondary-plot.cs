// Title: Aspose.Cells C# – Identify Secondary Plot Points in a Pie‑of‑Pie Chart (ChartPoint.IsInSecondaryPlot)
// Description: This example creates a workbook, adds category and value data, inserts a pie‑of‑pie chart, marks selected slices as secondary using ChartPoint.IsInSecondaryPlot, iterates the series to list points that belong to the secondary plot with their category names, and saves the file as PieChart_SecondaryPlot.xlsx.
// Keywords: Aspose.Cells ChartPoint.IsInSecondaryPlot | pie of pie secondary plot C# | Aspose.Cells secondary plot detection | .NET chart series secondary slice | identify secondary plot points Aspose | C# Aspose.Cells chart examples
// Common Searches: Aspose.Cells how to find secondary plot points | C# ChartPoint.IsInSecondaryPlot sample | pie of pie chart secondary slice Aspose.Cells | retrieve category of secondary plot points .NET | Aspose.Cells mark points as secondary plot
// Developer Intent: Find which data points in a pie‑of‑pie chart are assigned to the secondary plot using ChartPoint.IsInSecondaryPlot.
// Use Cases: Generate a report that lists categories of slices displayed in the secondary plot of a pie‑of‑pie chart. | Programmatically highlight specific data points by moving them to the secondary plot for clearer visualization. | Process or export secondary‑plot data separately after creating a combined chart in an automated workflow.
// AI Prompts: Write C# code with Aspose.Cells to create a pie‑of‑pie chart and output all points that have IsInSecondaryPlot set to true. | Explain the purpose of ChartPoint.IsInSecondaryPlot and demonstrate how to toggle it for selected series points. | Provide step‑by‑step instructions to read category labels of secondary plot points from a worksheet using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This example creates a workbook, adds category and value data, inserts a pie‑of‑pie chart, marks selected slices as secondary using ChartPoint.IsInSecondaryPlot, iterates the series to list points that belong to the secondary plot with their category names, and saves the file as PieChart_SecondaryPlot.xlsx.
    public class IdentifySecondaryPlotPoints
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (values and categories)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["A3"].PutValue("Orange");
                sheet.Cells["A4"].PutValue("Banana");
                sheet.Cells["A5"].PutValue("Grapes");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(50);
                sheet.Cells["B3"].PutValue(30);
                sheet.Cells["B4"].PutValue(15);
                sheet.Cells["B5"].PutValue(5);

                // Add a pie chart (pie of pie supports secondary plot)
                int chartIndex = sheet.Charts.Add(ChartType.Pie, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Mark some points as belonging to the secondary plot (last two points)
                chart.NSeries[0].Points[2].IsInSecondaryPlot = true; // Banana
                chart.NSeries[0].Points[3].IsInSecondaryPlot = true; // Grapes

                // Identify and output points that are in the secondary plot
                Console.WriteLine("Points in the secondary plot:");
                for (int i = 0; i < chart.NSeries[0].Points.Count; i++)
                {
                    ChartPoint point = chart.NSeries[0].Points[i];
                    if (point.IsInSecondaryPlot)
                    {
                        // Retrieve the category label for readability
                        string category = sheet.Cells[$"A{i + 2}"].StringValue;
                        Console.WriteLine($" - Index {i} (Category: {category})");
                    }
                }

                // Save the workbook
                string outputPath = "PieChart_SecondaryPlot.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            IdentifySecondaryPlotPoints.Run();
        }
    }
}
