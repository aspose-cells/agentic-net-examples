// Title: Log Secondary Plot Point Indices from an Aspose.Cells Pie Chart to a Text File (C#)
// Description: Creates a workbook, adds a pie chart with secondary plot points, detects points where IsInSecondaryPlot is true, writes their zero‑based indices to a .txt file, and saves the chart workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# chart API | secondary plot | IsInSecondaryPlot | log chart indices | write to text file | pie chart with secondary plot | .NET data visualization | chart analysis | Excel automation
// Common Searches: Aspose.Cells get secondary plot point indices | C# write chart point indexes to file | How to detect secondary plot points in Aspose.Cells | Save secondary plot indices from pie chart | Aspose.Cells chart logging example
// Developer Intent: Identify and persist the zero‑based indexes of chart points that belong to a secondary plot.
// Use Cases: Automated reporting of which data items are displayed in a secondary plot for business analytics. | Comparative analysis of secondary‑plot distribution across multiple workbooks. | Debugging and validation of chart rendering by recording secondary plot point positions.
// AI Prompts: Generate C# code using Aspose.Cells that extracts secondary plot point indexes from any chart type and saves them as CSV. | Show how to store the secondary plot indices from an Aspose.Cells chart into a SQL Server table with Entity Framework. | Provide a version of the example that logs the indices to the console and includes error handling for missing charts.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSecondaryPlotLogger
{
    // Creates a workbook, adds a pie chart with secondary plot points, detects points where IsInSecondaryPlot is true, writes their zero‑based indices to a .txt file, and saves the chart workbook using Aspose.Cells for .NET.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for a pie of pie chart
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

            // Mark some points as belonging to the secondary plot
            chart.NSeries[0].Points[2].IsInSecondaryPlot = true; // index 2
            chart.NSeries[0].Points[3].IsInSecondaryPlot = true; // index 3

            // Ensure chart calculations are up‑to‑date
            chart.Calculate();

            // Collect indices of points that are in the secondary plot
            var secondaryIndices = new System.Collections.Generic.List<int>();
            Series series = chart.NSeries[0];
            for (int i = 0; i < series.Points.Count; i++)
            {
                if (series.Points[i].IsInSecondaryPlot)
                {
                    secondaryIndices.Add(i);
                }
            }

            // Write the indices to a text file for later analysis
            string logFilePath = "SecondaryPlotPoints.txt";
            using (StreamWriter writer = new StreamWriter(logFilePath, false))
            {
                writer.WriteLine("Indices of secondary plot points:");
                foreach (int idx in secondaryIndices)
                {
                    writer.WriteLine(idx);
                }
            }

            // Save the workbook containing the chart
            workbook.Save("PieChartWithSecondaryPlot.xlsx");
        }
    }
}
