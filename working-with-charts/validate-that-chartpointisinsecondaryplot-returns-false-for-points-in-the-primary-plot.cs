using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsValidation
{
    class ValidateChartPointSecondaryPlot
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a pie chart (supports secondary plot)
            sheet.Cells["A1"].PutValue(50);
            sheet.Cells["A2"].PutValue(100);
            sheet.Cells["A3"].PutValue(150);
            sheet.Cells["B1"].PutValue(60);
            sheet.Cells["B2"].PutValue(32);
            sheet.Cells["B3"].PutValue(50);

            // Add a Pie chart
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Bind data to the chart
            chart.NSeries.Add("A1:B3", true);

            // Explicitly set a couple of points to be in the secondary plot
            // (these are the points we expect IsInSecondaryPlot == true)
            chart.NSeries[0].Points[2].IsInSecondaryPlot = true; // third point
            chart.NSeries[0].Points[3].IsInSecondaryPlot = true; // fourth point (the "Other" point)

            // Validate that all other points report false for IsInSecondaryPlot
            for (int i = 0; i < chart.NSeries[0].Points.Count; i++)
            {
                ChartPoint pt = chart.NSeries[0].Points[i];
                bool isSecondary = pt.IsInSecondaryPlot;

                // Points with index 2 and 3 were set to true; all others should be false
                bool expected = (i == 2 || i == 3);
                if (isSecondary != expected)
                {
                    Console.WriteLine($"Validation failed at point index {i}: Expected {expected}, got {isSecondary}");
                }
                else
                {
                    Console.WriteLine($"Point index {i}: IsInSecondaryPlot = {isSecondary} (as expected)");
                }
            }

            // Save the workbook (required by lifecycle rule)
            workbook.Save("ValidateIsInSecondaryPlot.xlsx");
        }
    }
}