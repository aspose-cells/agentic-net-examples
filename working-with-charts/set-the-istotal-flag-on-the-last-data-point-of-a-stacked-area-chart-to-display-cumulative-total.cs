using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class StackedAreaChartIsTotalDemo
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
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

            // Populate sample data for a stacked area chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");

            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a stacked area chart
            int chartIndex = sheet.Charts.Add(ChartType.AreaStacked, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Add series by column
            chart.NSeries.Add("=Sheet1!$B$2:$B$4", true);
            chart.NSeries.Add("=Sheet1!$C$2:$C$4", true);
            // Set category (X‑axis) data
            chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$4";

            // Determine the index of the last data point (zero‑based)
            int lastPointIndex = chart.NSeries[0].Points.Count - 1;

            // Mark the last point as a subtotal (cumulative total) for the first series
            chart.NSeries[0].LayoutProperties.Subtotals = new int[] { lastPointIndex };

            // Recalculate the chart to apply layout changes
            chart.Calculate();

            // Save the workbook
            string outputPath = "StackedAreaChartIsTotalDemo.xlsx";
            workbook.Save(outputPath);
        }
    }
}