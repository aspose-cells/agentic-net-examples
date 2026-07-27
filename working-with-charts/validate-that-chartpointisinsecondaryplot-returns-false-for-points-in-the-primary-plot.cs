using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartPointValidation
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a pie chart (3 categories)
            sheet.Cells["A1"].PutValue(50);
            sheet.Cells["A2"].PutValue(100);
            sheet.Cells["A3"].PutValue(150);
            sheet.Cells["B1"].PutValue(60);
            sheet.Cells["B2"].PutValue(32);
            sheet.Cells["B3"].PutValue(50);

            // Add a pie chart that supports secondary plots (Pie of Pie)
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("A1:B3", true);

            // Explicitly set the third point to be in the secondary plot
            // (indices are zero‑based)
            chart.NSeries[0].Points[2].IsInSecondaryPlot = true;

            // Validate that points not marked as secondary return false
            for (int i = 0; i < chart.NSeries[0].Points.Count; i++)
            {
                ChartPoint point = chart.NSeries[0].Points[i];
                bool isSecondary = point.IsInSecondaryPlot;

                if (i == 2) // this point was set to secondary
                {
                    if (!isSecondary)
                    {
                        Console.WriteLine($"Validation failed: Point {i} should be secondary but IsInSecondaryPlot is false.");
                    }
                    else
                    {
                        Console.WriteLine($"Point {i} correctly reports IsInSecondaryPlot = true.");
                    }
                }
                else // all other points should be primary
                {
                    if (isSecondary)
                    {
                        Console.WriteLine($"Validation failed: Point {i} should be primary but IsInSecondaryPlot is true.");
                    }
                    else
                    {
                        Console.WriteLine($"Point {i} correctly reports IsInSecondaryPlot = false.");
                    }
                }
            }

            // Save the workbook (output file name can be adjusted as needed)
            workbook.Save("ChartPointIsInSecondaryPlotValidation.xlsx");
        }
    }
}