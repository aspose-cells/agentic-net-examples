using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

class ProgressBarChartDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Prepare data: total value and current progress
            sheet.Cells["A1"].PutValue("Task");
            sheet.Cells["A2"].PutValue("Task 1");
            sheet.Cells["B1"].PutValue("Progress");
            sheet.Cells["B2"].PutValue(70);   // current progress
            sheet.Cells["C1"].PutValue("Total");
            sheet.Cells["C2"].PutValue(100);  // total length of the bar

            // Add a stacked column chart (used to simulate a progress bar)
            int chartIdx = sheet.Charts.Add(ChartType.ColumnStacked, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIdx];

            // Add the background series (the part that will be invisible)
            chart.NSeries.Add("C2:C2", true);          // total series
            chart.NSeries[0].Name = "Background";

            // Add the visible progress series
            chart.NSeries.Add("B2:B2", true);          // progress series
            chart.NSeries[1].Name = "Progress";

            // Set category (X‑axis) data
            chart.NSeries.CategoryData = "A2:A2";

            // Hide the background series (make it invisible)
            Series backgroundSeries = chart.NSeries[0];
            if (backgroundSeries.Area != null)
            {
                backgroundSeries.Area.Transparency = 1.0;          // 100% transparent fill
                if (backgroundSeries.Border != null)
                {
                    backgroundSeries.Border.Transparency = 1.0;   // hide border as well
                }
            }

            // Set a solid fill color for the visible progress series
            Series progressSeries = chart.NSeries[1];
            if (progressSeries.Area != null && progressSeries.Area.FillFormat != null)
            {
                progressSeries.Area.FillFormat.SolidFill.Color = Color.Green;
            }

            // Save the workbook
            workbook.Save("ProgressBarChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}