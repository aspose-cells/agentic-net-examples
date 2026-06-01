using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSecondaryPlotLogger
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a pie of pie chart
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
            chart.NSeries[0].Points[3].IsInSecondaryPlot = true; // index 3

            // Collect indices of points that are in the secondary plot
            List<string> secondaryIndices = new List<string>();
            Series series = chart.NSeries[0];
            for (int i = 0; i < series.Points.Count; i++)
            {
                if (series.Points[i].IsInSecondaryPlot)
                {
                    secondaryIndices.Add(i.ToString());
                }
            }

            // Write the indices to a text file for later analysis
            string outputPath = "SecondaryPlotIndices.txt";
            File.WriteAllLines(outputPath, secondaryIndices);

            // Save the workbook (optional, just to keep the chart)
            workbook.Save("PieChartWithSecondaryPlot.xlsx");
        }
    }
}