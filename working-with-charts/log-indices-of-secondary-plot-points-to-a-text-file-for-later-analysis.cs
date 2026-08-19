// Title: C# – Log Secondary Plot Point Indices from an Aspose.Cells Chart to a Text File
// Description: Creates a workbook, adds sample data, builds a pie‑of‑pie chart, marks selected points as secondary, extracts the indices of those points, and writes the list to a plain‑text file; the workbook can be saved optionally.
// Keywords: Aspose.Cells | C# | secondary plot | chart point index | pie of pie chart | IsInSecondaryPlot | export chart data to txt | log chart points | chart series iteration | write indices to file
// Common Searches: Aspose.Cells get secondary plot point indices | C# write chart point indices to text file | How to identify secondary plot points in Aspose.Cells | Export secondary plot data from pie‑of‑pie chart | Log secondary plot points Aspose.Cells C#
// Developer Intent: Find the points that belong to a chart's secondary plot and persist their indices for later processing or analysis.
// Use Cases: Debugging the distribution of data between primary and secondary sections of a pie‑of‑pie chart. | Feeding secondary‑plot indices into a downstream routine that applies custom formatting or calculations. | Collecting usage statistics of secondary plots across multiple generated workbooks. | Creating an audit trail of chart modifications for regulatory compliance.
// AI Prompts: Generate C# code using Aspose.Cells that returns a list of (seriesIndex, pointIndex) tuples for all secondary‑plot points and saves them as CSV. | Provide a reusable method that accepts a Chart object and returns a formatted string with series and point indices of secondary‑plot points. | Explain how to extend the sample to handle charts with multiple series and include both series and point indices in the log file.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSecondaryPlotLogger
{
    // Creates a workbook, adds sample data, builds a pie‑of‑pie chart, marks selected points as secondary, extracts the indices of those points, and writes the list to a plain‑text file; the workbook can be saved optionally.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data for a pie of pie chart
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
            chart.NSeries[0].Points[3].IsInSecondaryPlot = true; // index 3 (if exists)

            // Collect indices of points that are in the secondary plot
            var secondaryIndices = new System.Collections.Generic.List<int>();

            // Iterate through all series in the chart
            for (int s = 0; s < chart.NSeries.Count; s++)
            {
                Series series = chart.NSeries[s];
                // Iterate through points in the series
                for (int p = 0; p < series.Points.Count; p++)
                {
                    ChartPoint point = series.Points[p];
                    if (point.IsInSecondaryPlot)
                    {
                        // Store a composite index: series index and point index
                        // Here we store just the point index within its series
                        secondaryIndices.Add(p);
                    }
                }
            }

            // Write the collected indices to a text file
            string outputPath = "SecondaryPlotPoints.txt";
            using (StreamWriter writer = new StreamWriter(outputPath, false))
            {
                writer.WriteLine("Indices of points in secondary plot:");
                foreach (int idx in secondaryIndices)
                {
                    writer.WriteLine(idx);
                }
            }

            // Save the workbook (optional, to keep the chart)
            workbook.Save("PieChartWithSecondaryPlot.xlsx");
        }
    }
}
