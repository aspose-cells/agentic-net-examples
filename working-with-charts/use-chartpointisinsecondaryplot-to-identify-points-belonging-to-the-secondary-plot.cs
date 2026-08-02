using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class IdentifySecondaryPlotPoints
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a pie‑of‑pie chart
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

                // Add a pie chart that supports secondary plots
                int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Mark some points as belonging to the secondary plot
                // In a pie‑of‑pie chart, points with IsInSecondaryPlot = true appear in the second pie
                chart.NSeries[0].Points[2].IsInSecondaryPlot = true; // Orange
                chart.NSeries[0].Points[3].IsInSecondaryPlot = true; // Banana

                // Ensure chart calculations are up‑to‑date
                chart.Calculate();

                // Identify and list points that are in the secondary plot
                Console.WriteLine("Points in the secondary plot:");
                foreach (ChartPoint point in chart.NSeries[0].Points)
                {
                    if (point.IsInSecondaryPlot)
                    {
                        // Retrieve the category (X value) and the Y value for display
                        string category = point.XValue?.ToString() ?? "N/A";
                        double value = Convert.ToDouble(point.YValue);
                        Console.WriteLine($"- Category: {category}, Value: {value}");
                    }
                }

                // Save the workbook
                string outputPath = "PieOfPie_SecondaryPlot.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}