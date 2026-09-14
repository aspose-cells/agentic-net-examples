// Title: How to detect and list secondary‑plot points in a pie‑of‑pie chart using Aspose.Cells for .NET (C#)
// AI Prompts: Show C# code that sets ChartPoint.IsInSecondaryPlot for selected series points and then iterates through the series to output each secondary‑plot point’s category and value. | Generate a complete Aspose.Cells example that creates a pie‑of‑pie chart, marks specific points as secondary, reads those points via IsInSecondaryPlot, and saves the workbook.
// Common Searches: Aspose.Cells C# retrieve points marked as secondary plot in a pie of pie chart | How to use ChartPoint.IsInSecondaryPlot to filter data in Aspose.Cells | Example of iterating chart points to find secondary plot in Aspose.Cells workbook | C# Aspose.Cells pie chart secondary series extraction | List categories of secondary plot points in Excel using Aspose.Cells API
// Tags: ChartPoint.IsInSecondaryPlot usage Aspose.Cells | pie-of-pie secondary plot extraction C# | enumerate chart points Aspose.Cells | filter secondary plot data Excel C# | save workbook with secondary plot Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a workbook, adds a pie‑of‑pie chart, marks selected points as belonging to the secondary plot via IsInSecondaryPlot, iterates over all chart points to identify and print the category and value of secondary‑plot points, and saves the result to an Excel file.
    public class IdentifySecondaryPlotPoints
    {
        // Entry point for the example
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
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a pie‑of‑pie chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["A5"].PutValue("D");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(50);
            sheet.Cells["B3"].PutValue(120);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["B5"].PutValue(80);

            // Add a pie chart that supports secondary plots
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Bind the data range to the chart
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Mark some points to appear in the secondary plot (indices 1 and 3)
            chart.NSeries[0].Points[1].IsInSecondaryPlot = true;
            chart.NSeries[0].Points[3].IsInSecondaryPlot = true;

            // Iterate through all points and identify those in the secondary plot
            Console.WriteLine("Points belonging to the secondary plot:");
            for (int i = 0; i < chart.NSeries[0].Points.Count; i++)
            {
                ChartPoint point = chart.NSeries[0].Points[i];
                if (point.IsInSecondaryPlot)
                {
                    // Output the category name and value of the secondary plot point
                    string category = sheet.Cells[i + 2, 0].StringValue; // Column A (zero‑based)
                    double value = Convert.ToDouble(point.YValue);      // Ensure proper conversion
                    Console.WriteLine($"  Index: {i}, Category: {category}, Value: {value}");
                }
            }

            // Save the workbook to an Excel file
            string outputPath = "PieChart_SecondaryPlotPoints.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
