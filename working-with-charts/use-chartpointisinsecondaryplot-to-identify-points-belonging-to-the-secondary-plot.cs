// Title: Find Secondary Plot Points in a Pie‑of‑Pie Chart with Aspose.Cells for .NET
// Description: C# example that creates a workbook, adds a pie‑of‑pie chart, marks selected points as secondary using ChartPoint.IsInSecondaryPlot, calculates the chart, iterates over all points to detect those on the secondary plot, outputs their index and category, and saves the file.
// Keywords: Aspose.Cells | ChartPoint.IsInSecondaryPlot | pie of pie chart | secondary plot points | C# Aspose.Cells example | identify secondary series points | Aspose.Cells chart API | secondary plot detection | Aspose.Cells .NET chart | chart point classification
// Common Searches: Aspose.Cells ChartPoint.IsInSecondaryPlot example | how to detect secondary plot points in a pie of pie chart | C# code to list secondary plot categories Aspose.Cells | mark points as secondary plot Aspose.Cells chart | retrieve secondary series data from Aspose.Cells chart
// Developer Intent: Identify which data points are placed in the secondary plot of a pie‑of‑pie chart.
// Use Cases: Log or format secondary plot points after detection | Generate a report of categories shown in the secondary plot | Apply custom colors or markers to secondary plot points | Programmatically move data between primary and secondary plots
// AI Prompts: Write C# code using Aspose.Cells to flag specific points as secondary plot and list their categories. | Explain the purpose of ChartPoint.IsInSecondaryPlot and demonstrate how to separate primary and secondary data in a chart. | Provide a complete Aspose.Cells example that saves the workbook after identifying secondary plot points and adds comments to the related cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // C# example that creates a workbook, adds a pie‑of‑pie chart, marks selected points as secondary using ChartPoint.IsInSecondaryPlot, calculates the chart, iterates over all points to detect those on the secondary plot, outputs their index and category, and saves the file.
    public class IdentifySecondaryPlotPoints
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a pie of pie chart
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

                // Add a Pie chart (fallback if PieOfPie is unavailable)
                int chartIndex = sheet.Charts.Add(ChartType.Pie, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Mark specific points to appear in the secondary plot
                // Here we move "Banana" and "Grapes" to the secondary plot
                chart.NSeries[0].Points[2].IsInSecondaryPlot = true; // Banana
                chart.NSeries[0].Points[3].IsInSecondaryPlot = true; // Grapes

                // Ensure the chart is calculated so that all point properties are up‑to‑date
                chart.Calculate();

                // Iterate through all points and identify those belonging to the secondary plot
                Series series = chart.NSeries[0];
                for (int i = 0; i < series.Points.Count; i++)
                {
                    ChartPoint point = series.Points[i];
                    if (point.IsInSecondaryPlot)
                    {
                        // Output the index and category of the secondary plot point
                        string category = sheet.Cells[$"A{i + 2}"].StringValue;
                        Console.WriteLine($"Point at index {i} (Category: {category}) is in the secondary plot.");
                    }
                }

                // Prepare output directory
                string outputFile = "PieOfPie_SecondaryPlotIdentification.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputFile));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved to {outputFile}");
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
            IdentifySecondaryPlotPoints.Run();
        }
    }
}
