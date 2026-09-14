// Title: Log secondary plot point indices from an Aspose.Cells pie chart to a text file using C#
// AI Prompts: Generate C# code that creates a pie chart with Aspose.Cells, marks chosen points as secondary plot, iterates the series to collect their indices, and saves the list to a .txt file. | Show how to use Aspose.Cells ChartPoint.IsInSecondaryPlot property to filter secondary plot points and write both the index and the point value to a CSV file. | Adapt the example to log secondary plot point indices for multiple series in a workbook and store each series' results in separate text files.
// Common Searches: how to export secondary plot point indices from Aspose.Cells chart in C# | Aspose.Cells C# write chart point indexes to a text file | retrieve indices of secondary plot points in a pie chart using Aspose.Cells | C# Aspose.Cells example for logging secondary series points | save secondary plot data from Aspose.Cells workbook to txt
// Tags: Aspose.Cells secondary plot point logging | C# write chart point indices to txt | Aspose.Cells pie chart secondary plot extraction | iterate Aspose.Cells chart series points | log secondary plot data Aspose.Cells workbook

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSecondaryPlotLogger
{
    // Creates a workbook with a pie chart, marks specific points as secondary plot, collects their indices, writes the indices to a text file, and saves the workbook.
    class Program
    {
        static void Main()
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

            // Mark some points as belonging to the secondary plot
            chart.NSeries[0].Points[2].IsInSecondaryPlot = true; // index 2
            chart.NSeries[0].Points[3].IsInSecondaryPlot = true; // index 3 (if exists)

            // Ensure the chart is calculated so point data is up‑to‑date
            chart.Calculate();

            // Collect indices of points that are in the secondary plot
            List<int> secondaryIndices = new List<int>();
            Series series = chart.NSeries[0];
            for (int i = 0; i < series.Points.Count; i++)
            {
                ChartPoint point = series.Points[i];
                if (point.IsInSecondaryPlot)
                {
                    secondaryIndices.Add(i);
                }
            }

            // Prepare output lines
            List<string> outputLines = new List<string>();
            outputLines.Add("Indices of secondary plot points:");
            foreach (int idx in secondaryIndices)
            {
                outputLines.Add(idx.ToString());
            }

            // Write the indices to a text file for later analysis
            string outputPath = "SecondaryPlotPoints.txt";
            File.WriteAllLines(outputPath, outputLines);

            // Save the workbook (optional, just to keep the chart)
            workbook.Save("PieChartWithSecondaryPlot.xlsx");
        }
    }
}
